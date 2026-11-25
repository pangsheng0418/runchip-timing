/**
 * TUR (Taiwan Ultra Runners) 路跑晶片計時平台 v3.9 — 公開即時賽事、歷年個人成績與八百壯士英雄榜控制器
 * 1. 「即時賽事」：僅顯示進行中／籌備中活動（已結束活動不出現於下拉選單）
 * 2. 「歷年成績」：自動合併系統「已結束活動」＋後台「匯入歷年成績」供個人跨場次查詢
 * 3. 「八百壯士」：自動彙整所有名稱含「八百壯士」之活動，提供「當年度成績整合資訊」與「歷年英雄榜查詢」
 */
(function () {
  'use strict';

  const DB = window.RaceDB;
  const TE = window.TimingEngine;
  const Sync = window.CloudSync;
  const Cert = window.CertificateGenerator;

  const state = {
    gasUrl: '',
    events: [],
    allRegistrations: [],
    allScans: [],
    historyRecords: [],
    mode: 'live', // 'live' | 'history' | 'warriors800' | 'certificate'
    selectedEventId: '',
    desiredEventId: '',
    initialParamEvent: '',
    computed: null,
    computedCache: new Map(),
    catFilter: 'ALL',
    searchQuery: '',
    historySubMode: 'person', // 'person' | 'event'
    historyYear: 'ALL',
    historyEventKey: 'ALL',
    historyEventCat: 'ALL',
    historyEventQuery: '',
    historyQuery: '',
    w800Year: '',
    w800HeroYear: '', // 指定年度 (如 '2026') 或 'ALL' (歷年總累計)
    w800HeroQuery: '',
    w800Modal: {
      open: false,
      step: 'events', // 'events' | 'runners'
      selectedEventKey: '',
      runnerQuery: ''
    },
    cert: {
      style: 'official', // 'official' | 'custom'
      year: 'ALL',
      eventKey: 'ALL',
      query: '',
      selectedEntryKey: '',
      entriesCache: [],
      orientation: 'landscape',
      format: 'classic_gold',
      theme: 'ivory_gold',
      showSplits: true,
      cOrientation: 'landscape',
      cStyle: 'light_glass',
      cShape: 'rounded',
      cPos: 'left',
      cZoom: 1,
      cOffsetX: 0,
      cOffsetY: 0,
      cQuote: '',
      cShowLogo: true,
      cShowSeal: true,
      cShowSplits: true,
      bgImage: null,
      photoImage: null,
      photoBox: null
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

  function isEventFinished(ev) {
    return Boolean(ev && ev.status === 'finished');
  }

  function is800EventName(name) {
    return String(name || '').includes('八百壯士');
  }

  function getOpenEvents() {
    return state.events.filter((e) => !isEventFinished(e));
  }

  function getFinishedEvents() {
    return state.events
      .filter((e) => isEventFinished(e))
      .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  }

  function getEventYear(dateStr) {
    const m = String(dateStr || '').match(/^(\d{4})/);
    return m ? m[1] : '';
  }

  function parseDurationStrToMs(str) {
    const s = String(str || '').trim();
    if (!s) return null;
    const m = s.match(/^(\d+):(\d{1,2})(?::(\d{1,2})(?:\.(\d{1,3}))?)?$/);
    if (!m) return null;
    const h = parseInt(m[1], 10) || 0;
    const min = parseInt(m[2], 10) || 0;
    const sec = m[3] !== undefined ? (parseInt(m[3], 10) || 0) : 0;
    const ms = m[4] !== undefined ? parseInt(m[4].padEnd(3, '0'), 10) : 0;
    return ((h * 3600 + min * 60 + sec) * 1000) + ms;
  }

  function getComputedForEvent(evt) {
    if (!evt) return null;
    if (state.computedCache.has(evt.eventId)) {
      return state.computedCache.get(evt.eventId);
    }
    const eventRegs = state.allRegistrations.filter((r) => r.eventId === evt.eventId);
    const eventScans = state.allScans.filter((s) => s.eventId === evt.eventId && !s.deleted);
    const result = TE.computeEventResults({
      event: evt,
      runners: eventRegs,
      scans: eventScans
    });
    state.computedCache.set(evt.eventId, result);
    return result;
  }

  /**
   * 將「系統活動」與「匯入歷年成績」統一轉成標準個人參賽條目，便於合併查詢與八百壯士整合
   */
  function buildUnifiedEntries({ includeLive = false, includeFinished = true, includeImported = true, only800 = false } = {}) {
    const entries = [];

    // 1. 系統內活動（進行中／已結束）
    for (const ev of state.events) {
      const fin = isEventFinished(ev);
      if (fin && !includeFinished) continue;
      if (!fin && !includeLive) continue;
      if (only800 && !is800EventName(ev.name)) continue;

      const computed = getComputedForEvent(ev);
      if (!computed) continue;
      const isLap = ev.eventType === 'lap_timed';
      const evDate = ev.date || '';
      const year = getEventYear(evDate);
      const eventKey = `SYS__${ev.eventId}`;

      for (const row of computed.rows) {
        let completed = false;
        let distanceKm = 0;
        let finishTimeText = '';
        let finishTimeMs = null;
        let statusText = '未完賽';

        if (isLap) {
          completed = row.lapCount > 0;
          distanceKm = Number(row.totalDistanceKm || 0);
          finishTimeMs = row.totalElapsedMs;
          finishTimeText = row.totalElapsedMs !== null ? TE.formatDuration(row.totalElapsedMs, false) : '';
          statusText = completed ? `${row.lapCount} 圈` : (row.checkedIn ? '已報到' : '未報到');
        } else {
          completed = row.status === 'finished';
          const cat = (ev.categories || []).find((c) => c.id === row.categoryId);
          distanceKm = completed ? Number(cat?.distanceKm || 0) : 0;
          finishTimeMs = row.gunDurationMs;
          finishTimeText = row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, false) : '';
          statusText = completed ? '完賽' : (row.status === 'on_course' ? '賽道上' : (row.checkedIn ? '已報到' : '未完賽'));
        }

        entries.push({
          sourceType: fin ? 'system_finished' : 'system_live',
          eventKey,
          eventId: ev.eventId,
          year,
          eventDate: evDate,
          eventName: ev.name,
          eventType: ev.eventType || 'distance',
          categoryName: row.categoryName || row.categoryId || '',
          bib: row.bib || '',
          name: row.name || '',
          idMasked: row.idMasked || '',
          idHash: row.idHash || '',
          completed,
          statusText,
          finishTimeText,
          finishTimeMs,
          distanceKm,
          laps: isLap ? (row.lapCount || 0) : 0,
          overallRank: row.overallRank || null,
          categoryRank: row.categoryRank || null,
          avgPace: row.avgPace || '--',
          checkedIn: Boolean(row.checkedIn),
          checkInTimeMs: row.checkInTimeMs || null,
          systemRow: row,
          systemComputed: computed,
          note: ''
        });
      }
    }

    // 2. 後台匯入之以前歷年活動成績
    if (includeImported) {
      for (const h of state.historyRecords || []) {
        if (!h.eventName || !h.name) continue;
        if (only800 && !is800EventName(h.eventName)) continue;

        const evDate = h.eventDate || (h.year ? `${h.year}-01-01` : '');
        const year = h.year || getEventYear(evDate);
        const statusStr = String(h.status || '').trim() || '完賽';
        const isUnfinished = statusStr.includes('未') || statusStr.toUpperCase() === 'DNF' || statusStr.toUpperCase() === 'DNS';
        const km = Number(h.distanceKm || 0);
        const laps = Number(h.laps || 0);
        const completed = !isUnfinished && (statusStr.includes('完賽') || Boolean(h.finishTime) || km > 0 || laps > 0);
        const finishTimeMs = parseDurationStrToMs(h.finishTime);

        entries.push({
          sourceType: 'imported',
          eventKey: `IMP__${evDate}__${h.eventName}`,
          eventId: '',
          year,
          eventDate: evDate,
          eventName: h.eventName,
          eventType: laps > 0 ? 'lap_timed' : 'imported',
          categoryName: h.categoryName || '',
          bib: h.bib || '',
          name: h.name,
          idMasked: h.idMasked || '',
          idHash: h.idHash || '',
          completed,
          statusText: statusStr,
          finishTimeText: TE.normalizeTimeStr ? TE.normalizeTimeStr(h.finishTime) : String(h.finishTime || '').replace(/^(-?\d{1,3}:\d{2}:\d{2})\.\d+$/, '$1'),
          finishTimeMs,
          distanceKm: km,
          laps,
          overallRank: h.overallRank || null,
          categoryRank: h.categoryRank || null,
          avgPace: (finishTimeMs && km > 0) ? TE.formatPace(finishTimeMs, km) : '--',
          checkedIn: completed,
          checkInTimeMs: null,
          systemRow: null,
          systemComputed: null,
          note: h.note || ''
        });
      }
    }

    return entries;
  }

  /**
   * 依個人分組：優先以 idHash 合併；若某筆匯入歷史資料無身份證，且系統內同姓名僅對應 1 人，自動合併至同一人名下
   */
  function groupEntriesByPerson(entries) {
    const byHash = new Map();
    const noHash = [];

    for (const e of entries) {
      if (e.idHash) {
        if (!byHash.has(e.idHash)) {
          byHash.set(e.idHash, { key: e.idHash, name: e.name, idMasked: e.idMasked, idHash: e.idHash, records: [] });
        }
        const p = byHash.get(e.idHash);
        if (!p.idMasked && e.idMasked) p.idMasked = e.idMasked;
        p.records.push(e);
      } else {
        noHash.push(e);
      }
    }

    // 建立姓名 -> 符合的 idHash 集合，若某姓名唯一對應 1 位 idHash 則自動合併
    const nameToHashes = new Map();
    byHash.forEach((p, hash) => {
      const nm = p.name.trim();
      if (!nameToHashes.has(nm)) nameToHashes.set(nm, []);
      nameToHashes.get(nm).push(hash);
    });

    const byNameFallback = new Map();
    for (const e of noHash) {
      const nm = e.name.trim();
      const matchedHashes = nameToHashes.get(nm);
      if (matchedHashes && matchedHashes.length === 1) {
        byHash.get(matchedHashes[0]).records.push(e);
      } else {
        const key = `NAME__${nm}`;
        if (!byNameFallback.has(key)) {
          byNameFallback.set(key, { key, name: nm, idMasked: e.idMasked || '', idHash: '', records: [] });
        }
        byNameFallback.get(key).records.push(e);
      }
    }

    return [...byHash.values(), ...byNameFallback.values()];
  }

  function getHistoryEventsCatalog() {
    const map = new Map();
    for (const ev of getFinishedEvents()) {
      const key = `SYS__${ev.eventId}`;
      const regCount = state.allRegistrations.filter((r) => r.eventId === ev.eventId).length;
      map.set(key, {
        eventKey: key,
        eventDate: ev.date || '',
        year: getEventYear(ev.date),
        eventName: ev.name,
        eventType: ev.eventType || 'distance',
        sourceType: 'system_finished',
        count: regCount
      });
    }
    for (const h of state.historyRecords || []) {
      const evDate = h.eventDate || (h.year ? `${h.year}-01-01` : '');
      const key = `IMP__${evDate}__${h.eventName}`;
      if (!map.has(key)) {
        map.set(key, {
          eventKey: key,
          eventDate: evDate,
          year: h.year || getEventYear(evDate),
          eventName: h.eventName,
          eventType: h.laps > 0 ? 'lap_timed' : 'imported',
          sourceType: 'imported',
          count: 0
        });
      }
      map.get(key).count += 1;
    }
    return Array.from(map.values()).sort((a, b) => String(b.eventDate || '').localeCompare(String(a.eventDate || '')));
  }

  function get800EventsCatalog() {
    const map = new Map();
    for (const ev of state.events) {
      if (!is800EventName(ev.name)) continue;
      const key = `SYS__${ev.eventId}`;
      map.set(key, {
        eventKey: key,
        eventDate: ev.date || '',
        year: getEventYear(ev.date),
        eventName: ev.name,
        status: ev.status || 'active',
        sourceType: isEventFinished(ev) ? 'system_finished' : 'system_live'
      });
    }
    for (const h of state.historyRecords || []) {
      if (!is800EventName(h.eventName)) continue;
      const evDate = h.eventDate || (h.year ? `${h.year}-01-01` : '');
      const key = `IMP__${evDate}__${h.eventName}`;
      if (!map.has(key)) {
        map.set(key, {
          eventKey: key,
          eventDate: evDate,
          year: h.year || getEventYear(evDate),
          eventName: h.eventName,
          status: 'imported',
          sourceType: 'imported'
        });
      }
    }
    return Array.from(map.values()).sort((a, b) => String(b.eventDate || '').localeCompare(String(a.eventDate || '')));
  }

  const FONT_SCALE_KEY = 'tur_font_scale_v1';
  const FONT_SCALE_OPTIONS = ['sm', 'md', 'lg', 'xl'];
  const FONT_SCALE_LABELS = { sm: '小', md: '標準', lg: '大', xl: '特大' };
  const PUB_SELECTED_EVENT_KEY = 'tur_pub_selected_event_v1';

  function getSavedPublicEventId() {
    try {
      return (localStorage.getItem(PUB_SELECTED_EVENT_KEY) || '').trim();
    } catch {
      return '';
    }
  }

  function syncUrlEventParam(eventId) {
    try {
      if (!window.history || typeof window.history.replaceState !== 'function') return;
      const url = new URL(location.href);
      const id = String(eventId || '').trim();
      if (id) {
        url.searchParams.set('event', id);
      } else {
        url.searchParams.delete('event');
      }
      window.history.replaceState(null, '', url.pathname + url.search + url.hash);
    } catch {}
  }

  function persistSelectedEvent(eventId, { syncUrl = true } = {}) {
    const id = String(eventId || '').trim();
    state.desiredEventId = id;
    state.selectedEventId = id;
    try {
      if (id) {
        localStorage.setItem(PUB_SELECTED_EVENT_KEY, id);
      } else {
        localStorage.removeItem(PUB_SELECTED_EVENT_KEY);
      }
    } catch {}
    try {
      if (id) DB.saveDeviceSettings({ selectedEventId: id });
    } catch {}
    if (syncUrl) syncUrlEventParam(id);
  }

  function selectPublicEvent(eventId) {
    persistSelectedEvent(eventId, { syncUrl: true });
    state.catFilter = 'ALL';
    renderEventSelector();
    recomputeAndRender();
  }

  function getSavedFontScale() {
    try {
      const s = localStorage.getItem(FONT_SCALE_KEY) || 'md';
      return FONT_SCALE_OPTIONS.includes(s) ? s : 'md';
    } catch {
      return 'md';
    }
  }

  function applyPublicFontScale(scale, save = true) {
    const valid = FONT_SCALE_OPTIONS.includes(scale) ? scale : 'md';
    document.documentElement.setAttribute('data-font-scale', valid);
    if (save) {
      try { localStorage.setItem(FONT_SCALE_KEY, valid); } catch {}
    }
    const btn = $('btn-pub-font-cycle');
    if (btn) btn.textContent = `🔤 字體：${FONT_SCALE_LABELS[valid]}`;
  }

  async function initPublicPage() {
    applyPublicFontScale(getSavedFontScale(), false);
    const params = new URLSearchParams(location.search);
    const paramApi = (params.get('api') || '').trim();
    const paramEvent = (params.get('event') || '').trim();
    const paramMode = (params.get('mode') || '').trim();
    const cfgApi = (window.APP_CONFIG && window.APP_CONFIG.DEFAULT_GAS_URL ? String(window.APP_CONFIG.DEFAULT_GAS_URL).trim() : '');
    const saved = DB.getDeviceSettings();
    const savedPubEvent = getSavedPublicEventId();
    const initialEvent = paramEvent || savedPubEvent || (saved.selectedEventId || '');

    state.gasUrl = paramApi || cfgApi || (saved.gasUrl || '').trim();
    state.initialParamEvent = paramEvent;
    state.desiredEventId = initialEvent;
    state.selectedEventId = initialEvent;
    if (paramMode === 'history' || paramMode === 'warriors800' || paramMode === 'certificate') state.mode = paramMode;

    bindEvents();
    await loadData(false);
    setInterval(() => loadData(true), 20000);
  }

  function normalizeCloudPayload(d) {
    if (!d) return { events: [], registrations: [], scans: [], historyRecords: [], activeEventId: '', isLegacy: false };
    if (Array.isArray(d.events) && d.events.length > 0) {
      return {
        events: d.events,
        registrations: Array.isArray(d.registrations) ? d.registrations : [],
        scans: Array.isArray(d.scans) ? d.scans : [],
        historyRecords: Array.isArray(d.historyRecords) ? d.historyRecords : [],
        activeEventId: d.globalConfig?.activeEventId || d.events[0].eventId,
        isLegacy: false
      };
    }
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
        historyRecords: [],
        activeEventId: legacyId,
        isLegacy: true
      };
    }
    return {
      events: [],
      registrations: Array.isArray(d.registrations) ? d.registrations : [],
      scans: Array.isArray(d.scans) ? d.scans : [],
      historyRecords: Array.isArray(d.historyRecords) ? d.historyRecords : [],
      activeEventId: '',
      isLegacy: false
    };
  }

  function ensureValidSelectedEvent(preferredDefaultId = '', { isPreliminaryLocal = false } = {}) {
    const openEvents = getOpenEvents();

    if (state.initialParamEvent) {
      const paramEvt = state.events.find((e) => e.eventId === state.initialParamEvent);
      if (paramEvt) {
        state.initialParamEvent = '';
        if (isEventFinished(paramEvt)) {
          state.mode = 'history';
          state.historyEventKey = `SYS__${paramEvt.eventId}`;
        }
      } else if (!isPreliminaryLocal) {
        state.initialParamEvent = '';
      }
    }

    if (!openEvents.length) {
      state.selectedEventId = isPreliminaryLocal ? state.desiredEventId : '';
      return;
    }

    const targetId = state.desiredEventId || state.selectedEventId;
    if (targetId && openEvents.some((e) => e.eventId === targetId)) {
      state.selectedEventId = targetId;
      state.desiredEventId = targetId;
      return;
    }

    // 首次開啟先顯示本機快取時，若尚未含指定活動，僅暫時顯示預設活動、不覆蓋 desiredEventId，等雲端回傳立即切回指定活動
    if (isPreliminaryLocal && state.desiredEventId) {
      if (preferredDefaultId && openEvents.some((e) => e.eventId === preferredDefaultId)) {
        state.selectedEventId = preferredDefaultId;
      } else {
        state.selectedEventId = openEvents[0].eventId;
      }
      return;
    }

    const fallbackId = (preferredDefaultId && openEvents.some((e) => e.eventId === preferredDefaultId))
      ? preferredDefaultId
      : openEvents[0].eventId;
    state.selectedEventId = fallbackId;
    state.desiredEventId = fallbackId;
  }

  let _hasRenderedPublicOnce = false;

  async function renderPublicSnapshot({ isPreliminaryLocal = false } = {}) {
    renderModeUI();
    renderEventSelector({ isPreliminaryLocal });
    renderHistoryFilters();
    recomputeAndRender();
    if (state.mode === 'history') {
      await runHistorySearch(state.historyQuery);
    } else if (state.mode === 'warriors800') {
      await renderWarriors800Tab();
    } else if (state.mode === 'certificate') {
      await renderPublicCertificateTab();
    }
  }

  async function loadData(silent = false) {
    const badge = $('pub-source-badge');

    // 首次進入立即先用本機 IndexedDB 快取秒開渲染（<50ms），不讓使用者乾等雲端喚醒
    if (!_hasRenderedPublicOnce && state.gasUrl) {
      try {
        await loadFromLocalIndexedDB({ isPreliminaryLocal: true });
        if (state.events.length > 0 || state.historyRecords.length > 0) {
          badge.textContent = '本機快取・更新中...';
          await renderPublicSnapshot({ isPreliminaryLocal: true });
          _hasRenderedPublicOnce = true;
        } else if (!silent) {
          badge.textContent = '載入中...';
        }
      } catch {}
    } else if (!silent) {
      badge.textContent = '更新中...';
    }

    state.computedCache.clear();

    try {
      if (state.gasUrl) {
        const res = await Sync.fetchPublicData(state.gasUrl);
        const norm = normalizeCloudPayload(res.data);

        if (norm.events.length > 0 || norm.historyRecords.length > 0) {
          state.events = norm.events;
          state.allRegistrations = norm.registrations;
          state.allScans = norm.scans;
          state.historyRecords = norm.historyRecords;
          ensureValidSelectedEvent(norm.activeEventId, { isPreliminaryLocal: false });
          badge.textContent = norm.isLegacy ? '雲端（相容）' : '雲端同步';

          if (!norm.isLegacy && res.data) {
            DB.mergeCloudData(res.data).catch(() => {});
          }
        } else {
          await loadFromLocalIndexedDB({ isPreliminaryLocal: false });
          badge.textContent = state.events.length || state.historyRecords.length ? '本機快取' : '尚無資料';
        }
      } else {
        await loadFromLocalIndexedDB({ isPreliminaryLocal: false });
        badge.textContent = '本機預覽';
      }
    } catch (err) {
      console.warn('雲端讀取失敗，切換本機快取:', err);
      if (!_hasRenderedPublicOnce) await loadFromLocalIndexedDB({ isPreliminaryLocal: false });
      badge.textContent = '本機快取（離線）';
    }

    _hasRenderedPublicOnce = true;
    renderModeUI();
    renderEventSelector({ isPreliminaryLocal: false });
    renderHistoryFilters();
    recomputeAndRender();
    if (state.mode === 'history') {
      await runHistorySearch(state.historyQuery);
    } else if (state.mode === 'warriors800') {
      await renderWarriors800Tab();
    } else if (state.mode === 'certificate') {
      if (!silent) await renderPublicCertificateTab();
    }
    $('pub-last-updated').textContent = TE.formatClockSec(Date.now());
  }

  async function loadFromLocalIndexedDB({ isPreliminaryLocal = false } = {}) {
    await DB.init();
    const [platform, regs, scans, hist] = await Promise.all([
      DB.getPlatform(),
      DB.getAllRegistrations(),
      DB.getAllScans(),
      DB.getAllHistoryRecords()
    ]);

    state.events = platform.events || [];
    ensureValidSelectedEvent(platform.activeEventId || '', { isPreliminaryLocal });

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
    state.historyRecords = await Promise.all(
      (hist || []).map(async (h) => ({
        ...h,
        idMasked: h.idMasked || (h.idNumber ? TE.maskIdNumber(h.idNumber) : ''),
        idHash: h.idHash || (h.idNumber ? await TE.sha256Hex(h.idNumber) : '')
      }))
    );
  }

  function getActiveEvent() {
    const openEvents = getOpenEvents();
    return openEvents.find((e) => e.eventId === state.selectedEventId) || openEvents[0] || null;
  }

  function switchMode(mode) {
    state.mode = ['live', 'history', 'warriors800', 'certificate'].includes(mode) ? mode : 'live';
    renderModeUI();
    renderEventSelector();
    if (state.mode === 'history') {
      renderHistoryFilters();
      runHistorySearch(state.historyQuery);
    } else if (state.mode === 'warriors800') {
      renderWarriors800Tab();
    } else if (state.mode === 'certificate') {
      renderPublicCertificateTab();
    } else {
      recomputeAndRender();
    }
  }

  function renderModeUI() {
    const histEvtCount = getHistoryEventsCatalog().length;
    const w800EvtCount = get800EventsCatalog().length;
    if ($('pub-history-count-badge')) $('pub-history-count-badge').textContent = String(histEvtCount);
    if ($('pub-800-count-badge')) $('pub-800-count-badge').textContent = String(w800EvtCount);

    document.querySelectorAll('.pub-mode-btn').forEach((btn) => {
      const m = btn.dataset.pubMode;
      const active = m === state.mode;
      if (m === 'live') {
        btn.className = active
          ? 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-black bg-sky-500 text-slate-950 transition'
          : 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
      } else if (m === 'history') {
        btn.className = active
          ? 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-black bg-amber-500 text-slate-950 transition'
          : 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
      } else if (m === 'certificate') {
        btn.className = active
          ? 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-black bg-emerald-500 text-slate-950 transition shadow'
          : 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-bold text-emerald-300 hover:bg-slate-800 transition';
      } else {
        btn.className = active
          ? 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-black bg-gradient-to-r from-amber-500 to-orange-500 text-slate-950 transition shadow'
          : 'pub-mode-btn px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-bold text-amber-300 hover:bg-slate-800 transition';
      }
    });

    $('pub-view-live').classList.toggle('hidden', state.mode !== 'live');
    $('pub-view-history').classList.toggle('hidden', state.mode !== 'history');
    $('pub-view-warriors800').classList.toggle('hidden', state.mode !== 'warriors800');
    if ($('pub-view-certificate')) $('pub-view-certificate').classList.toggle('hidden', state.mode !== 'certificate');

    const openEvents = getOpenEvents();
    $('pub-live-selector-bar').classList.toggle('hidden', !(state.mode === 'live' && openEvents.length > 0));

    if (state.mode === 'history') {
      $('pub-race-name').textContent = '📜 歷年個人成績查詢';
      $('pub-race-meta').textContent = `已收錄 ${histEvtCount} 場已結束與歷年匯入活動`;
    } else if (state.mode === 'warriors800') {
      $('pub-race-name').textContent = '⚔️ TUR 八百壯士系統';
      $('pub-race-meta').textContent = `當年度成績整合 ＋ 歷年英雄榜（共 ${w800EvtCount} 場）`;
    } else if (state.mode === 'certificate') {
      $('pub-race-name').textContent = '🏅 官方完賽證書列印';
      $('pub-race-meta').textContent = '支援官方中英文對照樣式，或客製化上傳底圖與個人照片下載 PNG／PDF';
    }
  }

  function renderEventSelector({ isPreliminaryLocal = false } = {}) {
    const openEvents = getOpenEvents();
    const sel = $('pub-event-select');
    const wrap = $('pub-event-select-wrap');
    const pills = $('pub-event-pills');
    const bar = $('pub-live-selector-bar');

    if (openEvents.length === 0) {
      sel.innerHTML = '';
      if (wrap) wrap.classList.add('hidden');
      if (pills) {
        pills.innerHTML = '';
        pills.classList.add('hidden');
      }
      if (bar) bar.classList.add('hidden');
      return;
    }

    ensureValidSelectedEvent('', { isPreliminaryLocal });
    if (bar) bar.classList.toggle('hidden', state.mode !== 'live');
    if (wrap) wrap.classList.remove('hidden');

    const statusTag = { active: '進行中', upcoming: '籌備中' };
    sel.innerHTML = openEvents
      .map((ev) => {
        const typeTag = ev.eventType === 'lap_timed' ? `🔁繞圈(${ev.lapDistanceKm}K/圈)` : '🏃距離賽';
        return `<option value="${esc(ev.eventId)}" style="background-color:#0f172a;color:#f8fafc;">${typeTag} ${esc(ev.name)} (${esc(ev.date || '')}) [${statusTag[ev.status] || ev.status}]</option>`;
      })
      .join('');
    sel.value = state.selectedEventId;

    if (pills) {
      pills.classList.toggle('hidden', openEvents.length <= 1);
      pills.innerHTML = openEvents.length <= 1
        ? ''
        : openEvents
            .map((ev) => {
              const active = ev.eventId === state.selectedEventId;
              const icon = ev.eventType === 'lap_timed' ? '🔁' : '🏃';
              return `<button type="button" data-pill-event="${esc(ev.eventId)}"
                class="px-2.5 py-1 rounded-lg text-xs font-bold transition ${
                  active
                    ? 'bg-sky-500 text-slate-950 shadow'
                    : 'bg-slate-950 hover:bg-slate-800 border border-slate-700 text-slate-300'
                }">${icon} ${esc(ev.name)}</button>`;
            })
            .join('');
    }
  }

  function recomputeAndRender() {
    const evt = getActiveEvent();
    const hasOpen = Boolean(evt);
    $('pub-live-empty').classList.toggle('hidden', hasOpen);
    $('pub-live-content').classList.toggle('hidden', !hasOpen);

    if (!evt) {
      state.computed = null;
      if (state.mode === 'live') {
        const histCount = getHistoryEventsCatalog().length;
        $('pub-race-name').textContent = '即時成績查詢';
        $('pub-race-meta').textContent = histCount > 0
          ? `目前無進行中活動（已有 ${histCount} 場歷年活動可查）`
          : '目前尚無路跑活動';
      }
      return;
    }

    state.computed = getComputedForEvent(evt);

    if (state.mode === 'live') {
      const isLap = evt.eventType === 'lap_timed';
      const gunMs = evt.unifiedGunTimeMs || (evt.categories || []).find((c) => c.gunTimeMs)?.gunTimeMs || null;
      const gunText = gunMs ? ` ｜ 起跑 ${TE.formatDateTimeSec(gunMs)}` : '';
      $('pub-race-name').textContent = (isLap ? '🔁 ' : '') + evt.name;
      $('pub-race-meta').textContent = isLap
        ? `${evt.date || '--'}${evt.location ? ` ｜ ${evt.location}` : ''}${gunText} ｜ 單圈 ${evt.lapDistanceKm} km`
        : `${evt.date || '--'}${evt.location ? ` ｜ ${evt.location}` : ''}${gunText}`;
    }

    $('pub-lb-title').textContent = evt.eventType === 'lap_timed' ? '🏆 即時里程排行榜' : '🏆 成績排行榜';

    renderOverview(evt);
    renderLeaderboard(evt);
    if (state.searchQuery) runSearch(state.searchQuery);
  }

  function renderOverview(evt) {
    const s = state.computed.summary;
    const isLap = evt.eventType === 'lap_timed';

    if (isLap) {
      $('pub-kpi-label-1').textContent = '報名';
      $('pub-kpi-label-checkin').textContent = '已報到';
      $('pub-kpi-label-2').textContent = '總圈數';
      $('pub-kpi-label-3').textContent = '總里程 (km)';
      $('pub-kpi-label-4').textContent = '最高里程 (km)';

      $('pub-kpi-total').textContent = String(s.totalRunners);
      $('pub-kpi-checkin').textContent = `${s.checkedInCount ?? 0}`;
      $('pub-kpi-finished').textContent = `${s.totalLapsAll} 圈`;
      $('pub-kpi-oncourse').textContent = `${s.totalKmAll.toFixed(2)}`;
      $('pub-kpi-notstarted').textContent = `${s.topDistanceKm.toFixed(2)}`;

      $('pub-cat-card-title').textContent = '⏱️ 組別進度';
      $('pub-right-card-title').textContent = '🏅 領先榜 (TOP 6)';

      $('pub-category-stats').innerHTML = !s.byCategory.length
        ? '<div class="text-xs text-slate-500">尚無組別</div>'
        : s.byCategory
            .map((c) => {
              const remainText = c.gunTimeMs === null
                ? '尚未起跑'
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
                    <span>報到 <strong class="text-sky-400">${c.checkedIn ?? 0}</strong> / ${c.total} 人 ｜ 累積 ${c.sumLaps} 圈</span>
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
      $('pub-kpi-label-1').textContent = '報名';
      $('pub-kpi-label-checkin').textContent = '已報到';
      $('pub-kpi-label-2').textContent = '已完賽';
      $('pub-kpi-label-3').textContent = '賽道上';
      $('pub-kpi-label-4').textContent = '未報到';

      $('pub-kpi-total').textContent = String(s.totalRunners);
      $('pub-kpi-checkin').textContent = String(s.checkedInCount ?? 0);
      $('pub-kpi-finished').textContent = String(s.finishedCount);
      $('pub-kpi-oncourse').textContent = String(s.onCourseCount);
      $('pub-kpi-notstarted').textContent = String(s.notCheckedInCount ?? 0);

      $('pub-cat-card-title').textContent = '🏃 分組進度';
      $('pub-right-card-title').textContent = '📍 站點通過';

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
                      報到 <strong class="text-sky-400">${c.checkedIn ?? 0}</strong> /
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

  /* ========================================================================
   * 模式 1：即時賽事選手查詢
   * ======================================================================== */
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

    const matches = state.computed.rows.filter((row) =>
      String(row.name || '').toLowerCase().includes(qLower)
    );

    container.classList.remove('hidden');

    if (matches.length === 0) {
      const openIdSet = new Set(getOpenEvents().map((e) => e.eventId));
      const otherOpenIds = new Set();
      state.allRegistrations.forEach((r) => {
        const hit = String(r.name || '').toLowerCase().includes(qLower);
        if (hit && openIdSet.has(r.eventId) && r.eventId !== state.selectedEventId) {
          otherOpenIds.add(r.eventId);
        }
      });

      const allHist = buildUnifiedEntries({ includeLive: false, includeFinished: true, includeImported: true });
      const hasHist = allHist.some((e) =>
        String(e.name || '').toLowerCase().includes(qLower)
      );

      const buttons = [];
      otherOpenIds.forEach((eid) => {
        const ev = state.events.find((x) => x.eventId === eid);
        buttons.push(`<button type="button" data-switch-pub-event="${esc(eid)}" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-sky-600 hover:bg-sky-500 text-white">切換至「${esc(ev ? ev.name : eid)}」 →</button>`);
      });
      if (hasHist) {
        buttons.push(`<button type="button" data-open-history-query="${esc(q)}" class="px-3 py-1.5 rounded-lg text-xs font-black bg-amber-500 hover:bg-amber-400 text-slate-950">📜 查看「${esc(q)}」歷年成績 →</button>`);
      }

      container.innerHTML = `
        <div class="rounded-2xl bg-slate-950 border border-slate-800 p-5 text-center text-sm text-slate-400">
          本場活動查無符合「<strong class="text-white">${esc(q)}</strong>」的選手資料。
          ${buttons.length ? `<div class="mt-2.5 flex flex-wrap justify-center gap-2">${buttons.join('')}</div>` : ''}
        </div>
      `;
      return;
    }

    const isLap = state.computed.eventType === 'lap_timed';
    container.innerHTML = isLap
      ? matches.slice(0, 10).map((row) => renderLapSearchCard(row)).join('')
      : matches.slice(0, 10).map((row) => renderDistanceSearchCard(row)).join('');
  }

  function renderOtherEventChips(row) {
    const activeEvt = getActiveEvent();
    const openIdSet = new Set(getOpenEvents().map((e) => e.eventId));
    const otherOpenRegs = row.idHash
      ? state.allRegistrations.filter((r) => r.idHash === row.idHash && r.eventId !== state.selectedEventId && openIdSet.has(r.eventId))
      : [];

    const allHist = buildUnifiedEntries({ includeLive: false, includeFinished: true, includeImported: true });
    const histCount = allHist.filter((e) =>
      (row.idHash && e.idHash === row.idHash) || (!e.idHash && e.name === row.name)
    ).length;

    const all800 = buildUnifiedEntries({ includeLive: true, includeFinished: true, includeImported: true, only800: true });
    const w800Count = all800.filter((e) =>
      (row.idHash && e.idHash === row.idHash) || (!e.idHash && e.name === row.name)
    ).length;

    const canPrintLiveCert = Boolean(
      activeEvt && isEventFinished(activeEvt) && ((activeEvt.eventType === 'lap_timed' && Number(row.lapCount || 0) > 0) || row.status === 'finished')
    );

    if (!otherOpenRegs.length && !histCount && !w800Count && !canPrintLiveCert) return '';

    const certChip = canPrintLiveCert
      ? `<button type="button" data-open-cert-eventkey="SYS__${esc(activeEvt.eventId)}" data-open-cert-bib="${esc(row.bib)}" data-open-cert-name="${esc(row.name)}" class="px-2.5 py-0.5 rounded bg-emerald-500/20 hover:bg-emerald-500/35 border border-emerald-500/50 text-emerald-300 font-black">🏅 證書列印</button>`
      : '';

    const openChips = otherOpenRegs.map((or) => {
      const ev = state.events.find((e) => e.eventId === or.eventId);
      return `<button type="button" data-switch-pub-event="${esc(or.eventId)}" class="px-2 py-0.5 rounded bg-slate-800 hover:bg-sky-700 text-sky-300 font-bold">${esc(ev ? ev.name : or.eventId)} (${esc(or.categoryId)})</button>`;
    }).join('');

    const historyChip = histCount > 0
      ? `<button type="button" data-open-history-query="${esc(row.name)}" data-open-history-idhash="${esc(row.idHash || '')}" class="px-2.5 py-0.5 rounded bg-amber-500/20 hover:bg-amber-500/35 border border-amber-500/40 text-amber-300 font-bold">📜 歷年成績 (${histCount} 場)</button>`
      : '';

    const w800Chip = w800Count > 0
      ? `<button type="button" data-open-800-hero="${esc(row.name)}" class="px-2.5 py-0.5 rounded bg-orange-500/20 hover:bg-orange-500/35 border border-orange-500/40 text-orange-300 font-bold">⚔️ 八百壯士 (${w800Count} 場)</button>`
      : '';

    return `<div class="flex flex-wrap items-center gap-1.5 text-xs pt-1.5">${certChip}${openChips}${historyChip}${w800Chip}</div>`;
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
              單圈 ${row.lapDistanceKm} km
            </div>
            ${renderOtherEventChips(row)}
          </div>
          <div class="text-right">
            <div class="text-[11px] text-slate-400">最後通過時間</div>
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
            <div class="text-[11px] text-slate-400 font-bold">總名次 / 組名次</div>
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
          <div class="text-xs font-extrabold text-slate-300">🔁 逐圈明細</div>
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
      on_course: { text: '🏃 賽道上', cls: 'bg-amber-500/20 text-amber-300 border-amber-500/40' },
      checked_in: { text: '✅ 已報到', cls: 'bg-sky-500/20 text-sky-300 border-sky-500/40' },
      not_checked_in: { text: '⏳ 未報到', cls: 'bg-slate-800 text-slate-300 border-slate-700' }
    };
    const st = statusBanner[row.status] || statusBanner.not_checked_in;
    const lastLoc = row.lastCheckpoint
      ? `${row.lastCheckpoint.name}（${TE.formatClockSec(row.lastScan.timestampMs)}）`
      : (row.checkedIn ? '已報到' : '尚未報到');

    const etaBlock =
      row.status === 'on_course' && row.etaTimestampMs
        ? `<div class="bg-amber-950/40 border border-amber-500/40 rounded-xl p-3">
             <div class="text-[11px] text-amber-300 font-bold">預估完賽 (ETA)</div>
             <div class="text-lg font-black font-mono-num text-amber-300 mt-0.5">${TE.formatClockSec(row.etaTimestampMs)}</div>
             <div class="text-[11px] text-amber-200/80 font-mono-num">預估耗時 ${TE.formatDuration(row.etaDurationMs, false)}</div>
           </div>`
        : `<div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
             <div class="text-[11px] text-slate-400 font-bold">總名次 / 組名次</div>
             <div class="text-lg font-black font-mono-num text-emerald-400 mt-0.5">
               ${row.overallRank ? `總 #${row.overallRank}` : '--'} / ${row.categoryRank ? `組 #${row.categoryRank}` : '--'}
             </div>
             <div class="text-[11px] text-slate-400 font-mono-num">配速 ${esc(row.avgPace)}</div>
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
                ${split && split.durationMs !== null ? `<div class="text-[11px] font-mono-num text-sky-400">分段 ${TE.formatDuration(split.durationMs, false)} • 配速 ${esc(split.pace)}</div>` : ''}
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
              組別：<strong class="text-slate-200">${esc(row.categoryName)}</strong>
            </div>
            ${renderOtherEventChips(row)}
          </div>
          <div class="text-right">
            <div class="text-[11px] text-slate-400">最新動態</div>
            <div class="text-xs sm:text-sm font-bold text-sky-300">${esc(lastLoc)}</div>
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">起跑時間</div>
            <div class="text-base sm:text-lg font-black font-mono-num text-sky-300 mt-0.5">
              ${row.gunTimeMs ? TE.formatDateTimeSec(row.gunTimeMs) : '未設定'}
            </div>
            <div class="text-[11px] text-slate-400">
              狀態：${st.text}
            </div>
          </div>
          <div class="bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div class="text-[11px] text-slate-400 font-bold">完賽時間</div>
            <div class="text-lg font-black font-mono-num text-emerald-400 mt-0.5">
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, false) : '--:--:--'}
            </div>
          </div>
          ${etaBlock}
        </div>

        <div class="pt-1 space-y-2">
          <div class="text-xs font-extrabold text-slate-300">📍 各站分段紀錄</div>
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
          <th class="py-2.5 pr-3">限時組別</th>
          <th class="py-2.5 pr-3">圈數</th>
          <th class="py-2.5 pr-3">總里程</th>
          <th class="py-2.5 pr-3">耗時</th>
          <th class="py-2.5 pr-3">平均單圈</th>
          <th class="py-2.5">配速</th>
        </tr>
      `;
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="py-8 text-center text-slate-500">尚無繞圈成績資料</td></tr>';
        return;
      }
      tbody.innerHTML = rows
        .map((row) => `
          <tr class="hover:bg-slate-800/40 cursor-pointer" data-pick-name="${esc(row.name)}">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(row.name)}</td>
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
        <th class="py-2.5 pr-3">組別</th>
        <th class="py-2.5 pr-3">最新動態</th>
        <th class="py-2.5 pr-3">完賽時間</th>
        <th class="py-2.5">配速 / ETA</th>
      </tr>
    `;

    if (rows.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" class="py-8 text-center text-slate-500">尚無報名或成績資料</td></tr>';
      return;
    }

    tbody.innerHTML = rows
      .map((row) => {
        const locStr =
          row.status === 'finished'
            ? '<span class="text-emerald-400 font-bold">🏁 已完賽</span>'
            : row.lastCheckpoint
              ? `<span class="text-amber-300 font-bold">🏃 ${esc(row.lastCheckpoint.name)}</span>`
              : row.status === 'on_course'
                ? '<span class="text-amber-300 font-bold">🏃 賽道上</span>'
                : row.checkedIn
                  ? '<span class="text-sky-300 font-bold">✅ 已報到</span>'
                  : '<span class="text-slate-500">未報到</span>';

        const paceOrEta =
          row.status === 'finished'
            ? esc(row.avgPace)
            : row.etaTimestampMs
              ? `<span class="text-amber-300">預估 ${TE.formatClockSec(row.etaTimestampMs)}</span>`
              : '--';

        return `
          <tr class="hover:bg-slate-800/40 cursor-pointer" data-pick-name="${esc(row.name)}">
            <td class="py-2.5 pr-3 font-mono-num font-black text-amber-400">${row.overallRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-sky-300">${row.categoryRank ?? '-'}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-white">${esc(row.bib)}</td>
            <td class="py-2.5 pr-3 font-bold text-white">${esc(row.name)}</td>
            <td class="py-2.5 pr-3">
              <span class="inline-flex items-center gap-1.5 text-xs font-bold">
                <span class="w-2 h-2 rounded-full" style="background:${esc(row.categoryColor)}"></span>
                ${esc(row.categoryName)}
              </span>
            </td>
            <td class="py-2.5 pr-3">${locStr}</td>
            <td class="py-2.5 pr-3 font-mono-num font-bold text-emerald-400">
              ${row.gunDurationMs !== null ? TE.formatDuration(row.gunDurationMs, false) : '--'}
            </td>
            <td class="py-2.5 font-mono-num text-slate-300">${paceOrEta}</td>
          </tr>
        `;
      })
      .join('');
  }

  /* ========================================================================
   * 模式 2：歷年成績查詢（按年度區分：個人成績查詢 ／ 選取場次成績查詢）
   * ======================================================================== */
  function renderHistorySubModeUI() {
    const isEventMode = state.historySubMode === 'event';
    document.querySelectorAll('.pub-history-sub-btn').forEach((btn) => {
      const sub = btn.dataset.historySub;
      const active = sub === state.historySubMode;
      if (sub === 'person') {
        btn.className = active
          ? 'pub-history-sub-btn flex-1 px-3 py-2 rounded-lg text-xs sm:text-sm font-black bg-amber-500 text-slate-950 transition shadow'
          : 'pub-history-sub-btn flex-1 px-3 py-2 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
      } else {
        btn.className = active
          ? 'pub-history-sub-btn flex-1 px-3 py-2 rounded-lg text-xs sm:text-sm font-black bg-sky-500 text-slate-950 transition shadow'
          : 'pub-history-sub-btn flex-1 px-3 py-2 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
      }
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });

    $('form-pub-history-search')?.classList.toggle('hidden', isEventMode);
    $('pub-history-event-controls')?.classList.toggle('hidden', !isEventMode);

    const hasChosenEvent = isEventMode && state.historyEventKey !== 'ALL';
    const evtSel = $('pub-history-event-filter');
    if (evtSel) {
      evtSel.className = hasChosenEvent
        ? 'sm:col-span-6 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-xs sm:text-sm font-bold text-white focus:border-sky-400 focus:outline-none'
        : 'sm:col-span-12 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-xs sm:text-sm font-bold text-white focus:border-sky-400 focus:outline-none';
    }
    $('pub-history-event-cat-filter')?.classList.toggle('hidden', !hasChosenEvent);
    $('pub-history-event-search')?.classList.toggle('hidden', !hasChosenEvent);
  }

  function renderHistoryFilters() {
    const catalog = getHistoryEventsCatalog();
    const badge = $('pub-history-events-badge');

    const years = Array.from(new Set(catalog.map((c) => c.year).filter(Boolean))).sort((a, b) => b.localeCompare(a));
    const yearSel = $('pub-history-year-filter');
    yearSel.innerHTML = '<option value="ALL">全部年度</option>' + years.map((y) => `<option value="${esc(y)}">${esc(y)} 年度</option>`).join('');
    if (state.historyYear !== 'ALL' && !years.includes(state.historyYear)) {
      state.historyYear = 'ALL';
    }
    yearSel.value = state.historyYear;

    const filteredEvents = catalog.filter((c) => state.historyYear === 'ALL' || c.year === state.historyYear);
    if (badge) {
      badge.textContent = state.historyYear === 'ALL'
        ? `歷年 ${catalog.length} 場`
        : `${state.historyYear} 年度 ${filteredEvents.length} 場`;
    }

    const evtSel = $('pub-history-event-filter');
    if (evtSel) {
      evtSel.innerHTML =
        `<option value="ALL">-- 請選取${state.historyYear === 'ALL' ? '歷年' : `${state.historyYear} 年度`}場次（共 ${filteredEvents.length} 場） --</option>` +
        filteredEvents
          .map((c) => {
            const icon = is800EventName(c.eventName) ? '⚔️' : (c.eventType === 'lap_timed' ? '🔁' : '🏃');
            return `<option value="${esc(c.eventKey)}">${icon} ${esc(c.eventDate || '--')}｜${esc(c.eventName)} (${c.count}人)</option>`;
          })
          .join('');
      if (state.historyEventKey !== 'ALL' && !filteredEvents.some((c) => c.eventKey === state.historyEventKey)) {
        state.historyEventKey = 'ALL';
        state.historyEventCat = 'ALL';
        state.historyEventQuery = '';
        if ($('pub-history-event-search')) $('pub-history-event-search').value = '';
      }
      evtSel.value = state.historyEventKey;
    }

    renderHistorySubModeUI();
  }

  function renderHistoryEventCardsByYear(filteredCatalog, allHistEntries, hintText) {
    if (!filteredCatalog.length) {
      return `
        <div class="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center text-sm text-slate-400">
          此年度尚無歷年賽事紀錄。
        </div>
      `;
    }

    const entriesByEvent = new Map();
    for (const e of allHistEntries) {
      if (!entriesByEvent.has(e.eventKey)) entriesByEvent.set(e.eventKey, []);
      entriesByEvent.get(e.eventKey).push(e);
    }

    const byYear = new Map();
    for (const c of filteredCatalog) {
      const y = c.year || '未標註年度';
      if (!byYear.has(y)) byYear.set(y, []);
      byYear.get(y).push(c);
    }

    const yearsSorted = Array.from(byYear.keys()).sort((a, b) => b.localeCompare(a));
    const groupsHtml = yearsSorted.map((yr) => {
      const list = byYear.get(yr);
      const cards = list.map((c) => {
        const is800 = is800EventName(c.eventName);
        const icon = is800 ? '⚔️' : (c.eventType === 'lap_timed' ? '🔁' : '🏃');
        const rows = entriesByEvent.get(c.eventKey) || [];
        const finCount = rows.filter((r) => r.completed).length;
        const sumKm = rows.reduce((s, r) => s + Number(r.distanceKm || 0), 0);
        return `
          <button type="button" data-pick-history-event="${esc(c.eventKey)}"
            class="w-full text-left flex flex-wrap items-center justify-between gap-2.5 px-4 py-3 rounded-2xl bg-slate-950 border ${is800 ? 'border-amber-500/35 hover:border-amber-400' : 'border-slate-800 hover:border-sky-400'} hover:bg-slate-900 transition group">
            <div class="flex items-center gap-2 min-w-0">
              <span class="shrink-0">${icon}</span>
              <span class="font-mono-num text-xs font-bold text-amber-300 shrink-0">${esc(c.eventDate || '--')}</span>
              <span class="font-black text-white text-sm sm:text-base group-hover:text-sky-300 transition truncate">${esc(c.eventName)}</span>
              ${c.sourceType === 'imported' ? '<span class="px-1.5 py-0.5 rounded text-[10px] bg-slate-800 text-slate-400 shrink-0">歷年匯入</span>' : ''}
            </div>
            <div class="flex items-center gap-3 text-xs font-mono-num shrink-0 ml-auto">
              <span class="text-slate-400">參賽 <strong class="text-sky-300">${rows.length || c.count}</strong> ｜ 完賽 <strong class="text-emerald-400">${finCount}</strong>${sumKm > 0 ? ` ｜ ${sumKm.toFixed(0)} km` : ''}</span>
              <span class="font-black text-sky-400 group-hover:translate-x-0.5 transition">看成績 →</span>
            </div>
          </button>
        `;
      }).join('');

      return `
        <div class="space-y-2">
          <div class="flex items-center justify-between px-1">
            <span class="text-xs sm:text-sm font-black text-amber-300">📅 ${esc(yr)} 年度</span>
            <span class="text-xs font-mono-num text-slate-400">${list.length} 場</span>
          </div>
          <div class="space-y-2">${cards}</div>
        </div>
      `;
    }).join('');

    return `
      <div class="bg-slate-900 border border-slate-800 rounded-3xl p-4 sm:p-5 space-y-4">
        <div class="text-xs text-slate-400">${hintText}</div>
        <div class="space-y-4">${groupsHtml}</div>
      </div>
    `;
  }

  function renderHistoryEventView() {
    const container = $('pub-history-results');
    const catalog = getHistoryEventsCatalog();
    if (catalog.length === 0) {
      container.innerHTML = `
        <div class="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center space-y-1.5">
          <div class="text-3xl">📜</div>
          <div class="text-base font-black text-white">尚無歷年成績資料</div>
          <p class="text-xs text-slate-400">已結束之賽事或由後台匯入之以前歷年成績將自動合併顯示於此</p>
        </div>
      `;
      return;
    }

    const filteredCatalog = catalog.filter((c) => state.historyYear === 'ALL' || c.year === state.historyYear);
    const allHistEntries = buildUnifiedEntries({ includeLive: false, includeFinished: true, includeImported: true });

    renderHistorySubModeUI();

    // 尚未選取單一場次：按年度列出可點擊之場次清單
    if (state.historyEventKey === 'ALL') {
      const yrText = state.historyYear === 'ALL' ? '全部年度' : `${state.historyYear} 年度`;
      container.innerHTML = renderHistoryEventCardsByYear(
        filteredCatalog,
        allHistEntries,
        `目前顯示「${yrText}」共 ${filteredCatalog.length} 場歷年活動，請於上方選單選取或直接點選下方任一場次查看完整成績：`
      );
      return;
    }

    const targetEv = filteredCatalog.find((c) => c.eventKey === state.historyEventKey)
      || catalog.find((c) => c.eventKey === state.historyEventKey);
    if (!targetEv) {
      state.historyEventKey = 'ALL';
      renderHistoryFilters();
      renderHistoryEventView();
      return;
    }

    const evEntriesAll = allHistEntries
      .filter((e) => e.eventKey === targetEv.eventKey)
      .sort((a, b) => {
        const ra = a.overallRank ? Number(a.overallRank) : 999999;
        const rb = b.overallRank ? Number(b.overallRank) : 999999;
        if (ra !== rb) return ra - rb;
        if (a.completed !== b.completed) return a.completed ? -1 : 1;
        if (Number(b.distanceKm || 0) !== Number(a.distanceKm || 0)) return Number(b.distanceKm || 0) - Number(a.distanceKm || 0);
        if (a.finishTimeMs !== null && b.finishTimeMs !== null && a.finishTimeMs !== b.finishTimeMs) return a.finishTimeMs - b.finishTimeMs;
        return String(a.bib || '').localeCompare(String(b.bib || ''), 'zh-Hant', { numeric: true });
      });

    // 更新該場次組別下拉選單
    const cats = Array.from(new Set(evEntriesAll.map((e) => e.categoryName).filter(Boolean)));
    const catSel = $('pub-history-event-cat-filter');
    if (catSel) {
      catSel.innerHTML = '<option value="ALL">全部組別</option>' + cats.map((cat) => `<option value="${esc(cat)}">${esc(cat)}</option>`).join('');
      if (state.historyEventCat !== 'ALL' && !cats.includes(state.historyEventCat)) {
        state.historyEventCat = 'ALL';
      }
      catSel.value = state.historyEventCat;
    }

    const q = String(state.historyEventQuery || '').trim().toLowerCase();
    const shown = evEntriesAll.filter((e) => {
      if (state.historyEventCat !== 'ALL' && e.categoryName !== state.historyEventCat) return false;
      if (q) {
        return String(e.name || '').toLowerCase().includes(q);
      }
      return true;
    });

    const finCount = evEntriesAll.filter((e) => e.completed).length;
    const totalKm = evEntriesAll.reduce((s, e) => s + Number(e.distanceKm || 0), 0);
    const is800 = is800EventName(targetEv.eventName);
    const icon = is800 ? '⚔️' : (targetEv.eventType === 'lap_timed' ? '🔁' : '🏃');

    const tbodyHtml = !shown.length
      ? '<tr><td colspan="10" class="py-8 text-center text-slate-500">查無符合條件的選手成績</td></tr>'
      : shown.map((r, idx) => {
          const rankNum = r.overallRank || (r.completed ? idx + 1 : null);
          const rankDisp = rankNum === 1 ? '🥇' : rankNum === 2 ? '🥈' : rankNum === 3 ? '🥉' : (rankNum ? `#${rankNum}` : '-');
          const scoreText = r.laps > 0
            ? `${r.laps} 圈${r.finishTimeText ? ` (${r.finishTimeText})` : ''}`
            : (r.finishTimeText || '--');
          const kmText = Number(r.distanceKm || 0) > 0 ? `${Number(r.distanceKm).toFixed(2)} km` : '--';
          const statusCls = r.completed ? 'text-emerald-400 font-bold' : 'text-slate-400';
          return `
            <tr class="hover:bg-slate-800/45">
              <td class="py-2.5 px-3 font-mono-num font-black text-amber-400">${rankDisp}</td>
              <td class="py-2.5 px-2 font-mono-num font-bold text-sky-300">${r.categoryRank ? `#${r.categoryRank}` : '-'}</td>
              <td class="py-2.5 px-2 font-mono-num font-bold text-slate-200">${esc(r.bib || '--')}</td>
              <td class="py-2.5 px-2 font-bold text-white">
                <button type="button" data-open-history-query="${esc(r.name)}" data-open-history-idhash="${esc(r.idHash || '')}"
                  class="text-left font-bold text-white hover:text-amber-300 underline decoration-slate-600 hover:decoration-amber-400 underline-offset-2 transition">
                  ${esc(r.name)}
                </button>
              </td>
              <td class="py-2.5 px-2 text-sky-300 font-bold">${esc(r.categoryName || '--')}</td>
              <td class="py-2.5 px-2 font-mono-num font-black text-emerald-400">${esc(kmText)}</td>
              <td class="py-2.5 px-2 font-mono-num text-slate-200">${esc(scoreText)}</td>
              <td class="py-2.5 px-2 font-mono-num text-slate-300">${esc(r.avgPace || '--')}</td>
              <td class="py-2.5 px-2 ${statusCls}">${esc(r.statusText || (r.completed ? '完賽' : '未完賽'))}</td>
              <td class="py-2.5 px-3 text-right">
                ${r.completed
                  ? `<button type="button" data-open-cert-eventkey="${esc(r.eventKey)}" data-open-cert-bib="${esc(r.bib || '')}" data-open-cert-name="${esc(r.name)}" class="px-2 py-0.5 rounded-lg text-xs font-bold bg-emerald-500/20 hover:bg-emerald-500/35 border border-emerald-500/40 text-emerald-300">🏅 印證書</button>`
                  : '<span class="text-slate-600 text-xs">--</span>'}
              </td>
            </tr>
          `;
        }).join('');

    container.innerHTML = `
      <section class="bg-slate-900 border border-sky-500/40 rounded-3xl p-4 sm:p-5 space-y-4 shadow-xl">
        <div class="flex flex-wrap items-start justify-between gap-3 pb-3 border-b border-slate-800">
          <div class="space-y-1 min-w-0">
            <div class="flex items-center gap-2 flex-wrap">
              <button type="button" data-back-history-events="true"
                class="px-2.5 py-1 rounded-xl text-xs font-black bg-slate-800 hover:bg-slate-700 border border-slate-700 text-sky-300 transition">
                ← 返回場次清單
              </button>
              <span class="font-mono-num text-xs font-bold px-2 py-0.5 rounded bg-slate-800 text-amber-300">${esc(targetEv.eventDate || '--')}</span>
              ${is800 ? '<span class="px-2 py-0.5 rounded text-xs font-black bg-amber-500/20 text-amber-300">⚔️ 八百壯士</span>' : ''}
            </div>
            <h3 class="text-base sm:text-xl font-black text-white">${icon} ${esc(targetEv.eventName)}</h3>
          </div>
          <div class="flex flex-wrap items-center gap-2 text-xs font-mono-num">
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-slate-800 text-slate-300">參賽 <strong class="text-sky-300">${evEntriesAll.length}</strong> 人</span>
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-emerald-500/30 text-emerald-300">完賽 <strong>${finCount}</strong> 人</span>
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-amber-500/30 text-amber-300">總里程 <strong>${totalKm.toFixed(1)}</strong> km</span>
          </div>
        </div>

        <div class="flex items-center justify-between text-xs text-slate-400">
          <span>點選選手姓名可切換查詢該選手跨年度個人合併成績</span>
          <span class="font-mono-num">顯示 ${shown.length} / ${evEntriesAll.length} 筆</span>
        </div>

        <div class="overflow-x-auto border border-slate-800 rounded-2xl">
          <table class="w-full text-xs sm:text-sm text-left whitespace-nowrap">
            <thead class="bg-slate-950 text-xs text-slate-400 border-b border-slate-800">
              <tr>
                <th class="py-2.5 px-3">總名次</th>
                <th class="py-2.5 px-2">組名次</th>
                <th class="py-2.5 px-2">編號</th>
                <th class="py-2.5 px-2">姓名</th>
                <th class="py-2.5 px-2">組別</th>
                <th class="py-2.5 px-2">里程</th>
                <th class="py-2.5 px-2">成績 / 圈數</th>
                <th class="py-2.5 px-2">配速</th>
                <th class="py-2.5 px-2">狀態</th>
                <th class="py-2.5 px-3 text-right">證書</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-800/70">${tbodyHtml}</tbody>
          </table>
        </div>
      </section>
    `;
  }

  async function runHistorySearch(rawQuery, { forceIdHash = '' } = {}) {
    if (state.historySubMode === 'event' && !forceIdHash) {
      renderHistoryEventView();
      return;
    }

    const q = String(rawQuery ?? '').trim();
    state.historyQuery = q;
    const container = $('pub-history-results');
    const catalog = getHistoryEventsCatalog();

    if (catalog.length === 0) {
      container.innerHTML = `
        <div class="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center space-y-1.5">
          <div class="text-3xl">📜</div>
          <div class="text-base font-black text-white">尚無歷年成績資料</div>
          <p class="text-xs text-slate-400">已結束之賽事或由後台匯入之以前歷年成績將自動合併顯示於此</p>
        </div>
      `;
      return;
    }

    const filteredCatalog = catalog.filter((c) => state.historyYear === 'ALL' || c.year === state.historyYear);
    const allHistEntries = buildUnifiedEntries({ includeLive: false, includeFinished: true, includeImported: true });

    if (!q && !forceIdHash) {
      const yrLabel = state.historyYear === 'ALL' ? '全部年度' : `${state.historyYear} 年度`;
      container.innerHTML = renderHistoryEventCardsByYear(
        filteredCatalog,
        allHistEntries,
        `請於上方輸入「選手姓名」查詢「${yrLabel}」個人成績；或直接點選下方任一場次查看整場成績（共 ${filteredCatalog.length} 場）：`
      );
      return;
    }

    const qLower = q.toLowerCase();

    // 先取得該年度（或全部年度）歷年紀錄，依個人分組後僅以姓名篩選
    const scopedEntries = allHistEntries.filter((e) => state.historyYear === 'ALL' || e.year === state.historyYear);

    const people = groupEntriesByPerson(scopedEntries).filter((person) => {
      if (forceIdHash) return person.idHash === forceIdHash;
      return String(person.name || '').toLowerCase().includes(qLower);
    });

    if (people.length === 0) {
      const yrHint = state.historyYear === 'ALL' ? '歷年成績' : `${state.historyYear} 年度成績`;
      container.innerHTML = `
        <div class="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center text-sm text-slate-400">
          ${esc(yrHint)}中查無符合「<strong class="text-white">${esc(q)}</strong>」的個人紀錄。
        </div>
      `;
      return;
    }

    container.innerHTML = people.slice(0, 15).map((person) => renderPersonHistoryCard(person)).join('');
  }

  function renderUnifiedRecordRow(e) {
    const is800 = is800EventName(e.eventName);
    const isLap = e.eventType === 'lap_timed';
    const resultText = e.laps > 0
      ? `<span class="text-emerald-400 font-black">${e.distanceKm.toFixed(2)} km</span> <span class="text-purple-300 text-xs">(${e.laps} 圈)</span>`
      : (e.finishTimeText
          ? `<span class="text-emerald-400 font-black">${esc(e.finishTimeText)}</span>${e.distanceKm > 0 ? ` <span class="text-slate-400 text-xs">(${e.distanceKm} km)</span>` : ''}`
          : (e.distanceKm > 0
              ? `<span class="text-emerald-400 font-black">${e.distanceKm.toFixed(2)} km</span>`
              : `<span class="text-slate-500">${esc(e.statusText)}</span>`));

    const rankText = `${e.overallRank ? `總 #${e.overallRank}` : '--'} / ${e.categoryRank ? `組 #${e.categoryRank}` : '--'}`;

    let expandBlock = '';
    if (e.systemRow && e.systemComputed) {
      if (isLap) {
        const lapRows = e.systemRow.laps.length
          ? e.systemRow.laps.map((l) => `
              <tr class="${l.overtime ? 'text-rose-400/80' : 'text-slate-300'}">
                <td class="py-1 pr-3 font-mono-num">${l.overtime ? '超時' : `第 ${l.lapNumber} 圈`}</td>
                <td class="py-1 pr-3 font-mono-num text-sky-300">${TE.formatClockMs(l.timestampMs)}</td>
                <td class="py-1 pr-3 font-mono-num">${l.lapDurationMs !== null ? TE.formatDuration(l.lapDurationMs, false) : '--'}</td>
                <td class="py-1 pr-3 font-mono-num">${esc(l.lapPace)}</td>
                <td class="py-1 text-right font-mono-num text-emerald-400">${l.cumulativeKm !== null ? `${l.cumulativeKm.toFixed(2)} km` : '不計'}</td>
              </tr>
            `).join('')
          : '<tr><td colspan="5" class="py-2 text-center text-slate-500">無圈數紀錄</td></tr>';
        expandBlock = `
          <details class="pt-1 border-t border-slate-800/70 text-xs">
            <summary class="cursor-pointer font-bold text-sky-400 hover:text-sky-300 select-none">檢視逐圈明細 ▾</summary>
            <div class="overflow-x-auto pt-2">
              <table class="w-full text-xs text-left whitespace-nowrap">
                <thead class="text-slate-400 border-b border-slate-800">
                  <tr><th class="py-1 pr-3">圈次</th><th class="py-1 pr-3">時間</th><th class="py-1 pr-3">耗時</th><th class="py-1 pr-3">配速</th><th class="py-1 text-right">累積</th></tr>
                </thead>
                <tbody class="divide-y divide-slate-800/60">${lapRows}</tbody>
              </table>
            </div>
          </details>
        `;
      } else if ((e.systemComputed.checkpoints || []).length) {
        const cpRows = e.systemComputed.checkpoints.map((cp) => {
          const scan = e.systemRow.cpScanMap[cp.id];
          const split = e.systemRow.splits.find((sp) => sp.toId === cp.id);
          return `
            <tr class="text-slate-300">
              <td class="py-1 pr-3 font-bold">${esc(cp.name)}</td>
              <td class="py-1 pr-3 font-mono-num ${scan ? 'text-emerald-300' : 'text-slate-600'}">${scan ? TE.formatClockMs(scan.timestampMs) : '未通過'}</td>
              <td class="py-1 pr-3 font-mono-num text-sky-300">${split && split.durationMs !== null ? TE.formatDuration(split.durationMs, false) : '--'}</td>
              <td class="py-1 text-right font-mono-num text-slate-400">${split ? esc(split.pace) : '--'}</td>
            </tr>
          `;
        }).join('');
        expandBlock = `
          <details class="pt-1 border-t border-slate-800/70 text-xs">
            <summary class="cursor-pointer font-bold text-sky-400 hover:text-sky-300 select-none">檢視各站分段成績 ▾</summary>
            <div class="overflow-x-auto pt-2">
              <table class="w-full text-xs text-left whitespace-nowrap">
                <thead class="text-slate-400 border-b border-slate-800">
                  <tr><th class="py-1 pr-3">站點</th><th class="py-1 pr-3">通過時間</th><th class="py-1 pr-3">分段耗時</th><th class="py-1 text-right">配速</th></tr>
                </thead>
                <tbody class="divide-y divide-slate-800/60">${cpRows}</tbody>
              </table>
            </div>
          </details>
        `;
      }
    } else if (e.note) {
      expandBlock = `<div class="text-[11px] text-slate-400 pt-1 border-t border-slate-800/60">備註：${esc(e.note)}</div>`;
    }

    return `
      <div class="rounded-2xl bg-slate-950 border ${is800 ? 'border-amber-500/35' : 'border-slate-800'} p-3.5 sm:p-4 space-y-2">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div class="flex items-center gap-2 flex-wrap">
            <span class="font-mono-num text-xs font-bold px-2 py-0.5 rounded bg-slate-800 text-amber-300">${esc(e.eventDate || '--')}</span>
            <span class="font-black text-white text-sm sm:text-base">${esc(e.eventName)}</span>
            ${is800 ? '<span class="px-1.5 py-0.5 rounded text-[10px] font-black bg-amber-500/20 text-amber-300">⚔️ 八百壯士</span>' : ''}
            ${e.categoryName ? `<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/15 text-sky-300">${esc(e.categoryName)}</span>` : ''}
            ${e.bib ? `<span class="font-mono-num text-xs text-slate-400">#${esc(e.bib)}</span>` : ''}
            ${e.sourceType === 'imported' ? '<span class="px-1.5 py-0.5 rounded text-[10px] bg-slate-800 text-slate-400">歷年匯入</span>' : ''}
            ${e.sourceType === 'system_live' ? '<span class="px-1.5 py-0.5 rounded text-[10px] bg-emerald-500/20 text-emerald-300">進行中</span>' : ''}
          </div>
          <div class="font-mono-num text-sm sm:text-base">${resultText}</div>
        </div>

        <div class="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-400 font-mono-num">
          <span>名次：<strong class="text-amber-300">${rankText}</strong> ｜ 配速：<strong class="text-slate-200">${esc(e.avgPace)}</strong></span>
          <div class="flex items-center gap-2">
            <span>${esc(e.statusText)}</span>
            ${e.completed && e.sourceType !== 'system_live'
              ? `<button type="button" data-open-cert-eventkey="${esc(e.eventKey)}" data-open-cert-bib="${esc(e.bib || '')}" data-open-cert-name="${esc(e.name)}" class="px-2 py-0.5 rounded-lg bg-emerald-500/20 hover:bg-emerald-500/35 border border-emerald-500/40 text-emerald-300 font-sans font-bold">🏅 印證書</button>`
              : ''}
          </div>
        </div>

        ${expandBlock}
      </div>
    `;
  }

  function renderPersonHistoryCard(person) {
    const records = [...person.records].sort((a, b) =>
      String(b.eventDate || '').localeCompare(String(a.eventDate || ''))
    );

    const totalRaces = records.length;
    const completedRaces = records.filter((r) => r.completed).length;
    const totalKm = records.reduce((sum, r) => sum + Number(r.distanceKm || 0), 0);
    const has800 = records.some((r) => is800EventName(r.eventName));

    // 按年度分組呈現個人歷年成績
    const byYear = new Map();
    for (const r of records) {
      const y = r.year || '未標註年度';
      if (!byYear.has(y)) byYear.set(y, []);
      byYear.get(y).push(r);
    }
    const yearsSorted = Array.from(byYear.keys()).sort((a, b) => b.localeCompare(a));
    const groupedHtml = yearsSorted.map((yr) => {
      const yrRecs = byYear.get(yr);
      const yrKm = yrRecs.reduce((s, r) => s + Number(r.distanceKm || 0), 0);
      return `
        <div class="space-y-2">
          <div class="flex items-center justify-between px-1">
            <span class="text-xs font-black text-amber-300">📅 ${esc(yr)} 年度</span>
            <span class="text-xs font-mono-num text-slate-400">${yrRecs.length} 場${yrKm > 0 ? ` ｜ ${yrKm.toFixed(2)} km` : ''}</span>
          </div>
          <div class="space-y-2.5">${yrRecs.map(renderUnifiedRecordRow).join('')}</div>
        </div>
      `;
    }).join('');

    return `
      <section class="bg-slate-900 border border-amber-500/30 rounded-3xl p-4 sm:p-5 space-y-4 shadow-xl">
        <div class="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-800">
          <div class="flex items-center gap-2.5 flex-wrap">
            <span class="text-xl sm:text-2xl font-black text-white">${esc(person.name)}</span>
            ${has800 ? `<button type="button" data-open-800-hero="${esc(person.name)}" class="px-2.5 py-0.5 rounded-full text-xs font-black bg-amber-500/20 hover:bg-amber-500/35 border border-amber-500/40 text-amber-300">⚔️ 查看八百壯士英雄榜</button>` : ''}
          </div>
          <div class="flex flex-wrap items-center gap-2 text-xs font-mono-num">
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-slate-800 text-slate-300">參賽 <strong class="text-white">${totalRaces}</strong> 場</span>
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-emerald-500/30 text-emerald-300">完賽 <strong>${completedRaces}</strong> 場</span>
            <span class="px-2.5 py-1 rounded-xl bg-slate-950 border border-amber-500/30 text-amber-300">累計 <strong>${totalKm.toFixed(2)}</strong> km</span>
          </div>
        </div>

        <div class="space-y-4">${groupedHtml}</div>
      </section>
    `;
  }

  /* ========================================================================
   * 模式 3：⚔️ TUR 八百壯士系統（當年度成績整合資訊 ＋ 歷年英雄榜查詢）
   * ======================================================================== */
  async function renderWarriors800Tab() {
    const all800Entries = buildUnifiedEntries({ includeLive: true, includeFinished: true, includeImported: true, only800: true });
    const catalog800 = get800EventsCatalog();
    const hasAny = catalog800.length > 0;

    $('pub-800-empty').classList.toggle('hidden', hasAny);
    $('pub-800-content').classList.toggle('hidden', !hasAny);
    if (!hasAny) return;

    // 1. 建立年度選單（預設當年度，若當年度無資料則取最新年度）
    const years = Array.from(new Set(catalog800.map((c) => c.year).filter(Boolean))).sort((a, b) => b.localeCompare(a));
    const currentYear = String(new Date().getFullYear());
    if (!state.w800Year || !years.includes(state.w800Year)) {
      state.w800Year = years.includes(currentYear) ? currentYear : (years[0] || currentYear);
    }
    if (!state.w800HeroYear || !years.includes(state.w800HeroYear)) {
      state.w800HeroYear = state.w800Year;
    }

    const yearSel = $('pub-800-year-select');
    yearSel.innerHTML = years.map((y) => `<option value="${esc(y)}">${esc(y)} 年度</option>`).join('');
    yearSel.value = state.w800Year;

    // 建立英雄榜年度選單與快速切換膠囊（僅列單一年度，註明當年度，不混合跨年度）
    const heroYearSel = $('pub-800-hero-year-select');
    if (heroYearSel) {
      heroYearSel.innerHTML = years
        .map((y) => `<option value="${esc(y)}">${esc(y)} 年度${y === currentYear ? '（當年度）' : ''}</option>`)
        .join('');
      heroYearSel.value = state.w800HeroYear;
    }
    const heroPills = $('pub-800-hero-year-pills');
    if (heroPills) {
      heroPills.innerHTML = years
        .map((y) => {
          const active = y === state.w800HeroYear;
          const label = y === currentYear ? `📅 ${y} 年度（當年度）` : `📅 ${y} 年度`;
          return `<button type="button" data-w800-hero-year="${esc(y)}" class="px-3 py-1 rounded-xl text-xs font-black transition ${active ? 'bg-amber-500 text-slate-950 shadow' : 'bg-slate-950 border border-slate-700 text-slate-300 hover:border-amber-400/70 hover:text-amber-200'}">${esc(label)}</button>`;
        })
        .join('');
    }

    // 2. 渲染「當年度八百壯士成績整合資訊」（僅顯示年度 KPI；場次與選手明細由「年度場次」數字點入鑽取查看）
    const yrCatalog = catalog800.filter((c) => c.year === state.w800Year);
    const yrEntries = all800Entries.filter((e) => e.year === state.w800Year);
    const yrFinishers = yrEntries.filter((e) => e.completed).length;
    const yrTotalKm = yrEntries.reduce((sum, e) => sum + Number(e.distanceKm || 0), 0);

    $('pub-800-yr-events').textContent = String(yrCatalog.length);
    $('pub-800-yr-entries').textContent = String(yrEntries.length);
    $('pub-800-yr-finishers').textContent = String(yrFinishers);
    $('pub-800-yr-km').textContent = yrTotalKm.toFixed(1);

    if (state.w800Modal.open) {
      renderWarriors800EventsModal(all800Entries, catalog800);
    }

    // 3. 渲染「⚔️ 八百壯士英雄榜（依選定年度區分顯示）」
    await renderWarriors800HeroList(all800Entries);
  }

  function openWarriors800EventsModal() {
    state.w800Modal.open = true;
    state.w800Modal.step = 'events';
    state.w800Modal.selectedEventKey = '';
    state.w800Modal.runnerQuery = '';
    if ($('modal-800-runner-search')) $('modal-800-runner-search').value = '';
    $('modal-800-events')?.classList.remove('hidden');
    renderWarriors800EventsModal();
  }

  function closeWarriors800EventsModal() {
    state.w800Modal.open = false;
    $('modal-800-events')?.classList.add('hidden');
  }

  function renderWarriors800EventsModal(preloadedEntries = null, preloadedCatalog = null) {
    const modal = $('modal-800-events');
    if (!modal || !state.w800Modal.open) return;

    const all800Entries = preloadedEntries || buildUnifiedEntries({ includeLive: true, includeFinished: true, includeImported: true, only800: true });
    const catalog800 = preloadedCatalog || get800EventsCatalog();
    const yrCatalog = catalog800.filter((c) => c.year === state.w800Year);
    const yrEntries = all800Entries.filter((e) => e.year === state.w800Year);

    const statusBadgeMap = {
      active: '<span class="px-2 py-0.5 rounded text-xs font-bold bg-emerald-500/20 text-emerald-300">進行中</span>',
      upcoming: '<span class="px-2 py-0.5 rounded text-xs font-bold bg-sky-500/20 text-sky-300">籌備中</span>',
      finished: '<span class="px-2 py-0.5 rounded text-xs font-bold bg-slate-800 text-slate-300">已結束</span>',
      imported: '<span class="px-2 py-0.5 rounded text-xs font-bold bg-amber-500/15 text-amber-300">歷年匯入</span>'
    };

    const stepEventsEl = $('modal-800-step-events');
    const stepRunnersEl = $('modal-800-step-runners');
    const backBtn = $('btn-modal-800-back');

    if (state.w800Modal.step === 'runners' && state.w800Modal.selectedEventKey) {
      const targetEv = yrCatalog.find((c) => c.eventKey === state.w800Modal.selectedEventKey);
      if (!targetEv) {
        state.w800Modal.step = 'events';
        state.w800Modal.selectedEventKey = '';
      } else {
        const evRowsAll = yrEntries
          .filter((e) => e.eventKey === targetEv.eventKey)
          .sort((a, b) => {
            const ra = a.overallRank ? Number(a.overallRank) : 999999;
            const rb = b.overallRank ? Number(b.overallRank) : 999999;
            if (ra !== rb) return ra - rb;
            if (a.completed !== b.completed) return a.completed ? -1 : 1;
            if (Number(b.distanceKm || 0) !== Number(a.distanceKm || 0)) return Number(b.distanceKm || 0) - Number(a.distanceKm || 0);
            if (a.finishTimeMs !== null && b.finishTimeMs !== null && a.finishTimeMs !== b.finishTimeMs) return a.finishTimeMs - b.finishTimeMs;
            return String(a.bib || '').localeCompare(String(b.bib || ''), 'zh-Hant', { numeric: true });
          });

        const evFin = evRowsAll.filter((r) => r.completed).length;
        const evKm = evRowsAll.reduce((s, r) => s + Number(r.distanceKm || 0), 0);

        $('modal-800-title').textContent = `⚔️ ${targetEv.eventName}`;
        $('modal-800-subtitle').textContent = `${targetEv.eventDate || '--'} ｜ 參賽 ${evRowsAll.length} 人 ｜ 完賽 ${evFin} 人 ｜ 總里程 ${evKm.toFixed(1)} km`;
        backBtn?.classList.remove('hidden');
        stepEventsEl?.classList.add('hidden');
        stepRunnersEl?.classList.remove('hidden');

        const q = String(state.w800Modal.runnerQuery || '').trim().toLowerCase();
        const filteredRows = q
          ? evRowsAll.filter((r) =>
              String(r.name || '').toLowerCase().includes(q)
              || String(r.bib || '').toLowerCase().includes(q)
              || String(r.categoryName || '').toLowerCase().includes(q)
            )
          : evRowsAll;

        if ($('modal-800-runner-count')) {
          $('modal-800-runner-count').textContent = q
            ? `符合 ${filteredRows.length} / ${evRowsAll.length} 人`
            : `共 ${evRowsAll.length} 人`;
        }

        const tbody = $('modal-800-runners-tbody');
        if (!filteredRows.length) {
          tbody.innerHTML = '<tr><td colspan="7" class="py-8 text-center text-slate-500">查無符合條件的選手成績明細</td></tr>';
        } else {
          tbody.innerHTML = filteredRows
            .map((r, idx) => {
              const rankNum = r.overallRank || (r.completed ? idx + 1 : null);
              const rankDisp = rankNum === 1 ? '🥇' : rankNum === 2 ? '🥈' : rankNum === 3 ? '🥉' : (rankNum ? `#${rankNum}` : '-');
              const scoreText = r.laps > 0
                ? `${r.laps} 圈${r.finishTimeText ? ` (${r.finishTimeText})` : ''}`
                : (r.finishTimeText || '--');
              const kmText = Number(r.distanceKm || 0) > 0 ? `${Number(r.distanceKm).toFixed(2)} km` : '--';
              const statusCls = r.completed ? 'text-emerald-400 font-bold' : 'text-slate-400';
              return `
                <tr class="hover:bg-slate-800/45">
                  <td class="py-2.5 px-3 font-mono-num font-black text-amber-400">${rankDisp}</td>
                  <td class="py-2.5 px-2 font-mono-num font-bold text-slate-200">${esc(r.bib || '--')}</td>
                  <td class="py-2.5 px-2 font-bold text-white">${esc(r.name)}</td>
                  <td class="py-2.5 px-2 text-sky-300 font-bold">${esc(r.categoryName || '--')}</td>
                  <td class="py-2.5 px-2 font-mono-num font-black text-emerald-400">${esc(kmText)}</td>
                  <td class="py-2.5 px-2 font-mono-num text-slate-200">${esc(scoreText)}</td>
                  <td class="py-2.5 px-3 ${statusCls}">${esc(r.statusText || (r.completed ? '完賽' : '未完賽'))}</td>
                </tr>
              `;
            })
            .join('');
        }
        return;
      }
    }

    // 第一層：列出當年度各場次清單（點任一場次進入第二層）
    $('modal-800-title').textContent = `⚔️ ${state.w800Year} 年度八百壯士場次明細`;
    $('modal-800-subtitle').textContent = `共 ${yrCatalog.length} 場 ｜ 點選任一場次查看所有選手明細成績`;
    backBtn?.classList.add('hidden');
    stepEventsEl?.classList.remove('hidden');
    stepRunnersEl?.classList.add('hidden');

    if (!yrCatalog.length) {
      stepEventsEl.innerHTML = '<div class="py-8 text-center text-slate-400">本年度尚無八百壯士場次資料</div>';
      return;
    }

    stepEventsEl.innerHTML = yrCatalog
      .map((ev, idx) => {
        const evRows = yrEntries.filter((e) => e.eventKey === ev.eventKey);
        const evFin = evRows.filter((e) => e.completed).length;
        const evKm = evRows.reduce((s, e) => s + Number(e.distanceKm || 0), 0);
        return `
          <button type="button" data-w800-event-key="${esc(ev.eventKey)}"
            class="w-full text-left p-4 rounded-2xl bg-slate-950 border border-slate-800 hover:border-amber-400/70 hover:bg-slate-900 transition space-y-2 group">
            <div class="flex items-center justify-between gap-2">
              <span class="font-black text-white text-sm sm:text-base group-hover:text-amber-300 transition">
                ${idx + 1}. ⚔️ ${esc(ev.eventName)}
              </span>
              <div class="flex items-center gap-2 shrink-0">
                ${statusBadgeMap[ev.status] || ''}
                <span class="text-xs font-black text-sky-400 group-hover:translate-x-0.5 transition">選手明細 →</span>
              </div>
            </div>
            <div class="flex flex-wrap items-center justify-between gap-2 text-xs font-mono-num text-slate-400">
              <span>📅 ${esc(ev.eventDate || '--')}</span>
              <span>參賽 <strong class="text-sky-300">${evRows.length}</strong> 人 ｜ 完賽 <strong class="text-emerald-400">${evFin}</strong> 人 ｜ 總里程 <strong class="text-amber-300">${evKm.toFixed(1)}</strong> km</span>
            </div>
          </button>
        `;
      })
      .join('');
  }

  async function renderWarriors800HeroList(preloadedEntries = null) {
    const all800Entries = preloadedEntries || buildUnifiedEntries({ includeLive: true, includeFinished: true, includeImported: true, only800: true });
    const currentYear = String(new Date().getFullYear());
    const heroYear = state.w800HeroYear || state.w800Year || currentYear;
    const isCurrentYear = heroYear === currentYear;
    const yearTag = isCurrentYear ? `${heroYear} 當年度` : `${heroYear} 年度`;
    const scopedEntries = all800Entries.filter((e) => e.year === heroYear);

    const allHeroes = groupEntriesByPerson(scopedEntries).map((p) => {
      const recs = [...p.records].sort((a, b) => String(b.eventDate || '').localeCompare(String(a.eventDate || '')));
      const totalKm = recs.reduce((s, r) => s + Number(r.distanceKm || 0), 0);
      const completed = recs.filter((r) => r.completed).length;
      const bestSingleKm = recs.reduce((m, r) => Math.max(m, Number(r.distanceKm || 0)), 0);
      const bestTimeMs = recs.reduce((best, r) => {
        if (!r.completed || r.finishTimeMs === null) return best;
        return best === null ? r.finishTimeMs : Math.min(best, r.finishTimeMs);
      }, null);
      return { ...p, records: recs, totalKm, completed, bestSingleKm, bestTimeMs };
    }).sort((a, b) => {
      if (b.totalKm !== a.totalKm) return b.totalKm - a.totalKm;
      if (b.completed !== a.completed) return b.completed - a.completed;
      if (a.bestTimeMs !== null && b.bestTimeMs !== null) return a.bestTimeMs - b.bestTimeMs;
      return a.name.localeCompare(b.name, 'zh-Hant');
    }).map((h, idx) => ({ ...h, heroRank: idx + 1 }));

    const totalAllKm = allHeroes.reduce((s, h) => s + h.totalKm, 0);
    if ($('pub-800-hero-title')) {
      $('pub-800-hero-title').textContent = `⚔️ ${yearTag}八百壯士英雄榜`;
    }
    $('pub-800-hero-summary').textContent = `${yearTag}共 ${allHeroes.length} 位英雄 ｜ 年度總里程 ${totalAllKm.toFixed(1)} km`;

    const q = (state.w800HeroQuery || '').trim();
    let shown = allHeroes;
    if (q) {
      const qLower = q.toLowerCase();
      shown = allHeroes.filter((h) =>
        String(h.name || '').toLowerCase().includes(qLower)
      );
    }

    const listEl = $('pub-800-hero-list');
    if (!shown.length) {
      listEl.innerHTML = q
        ? `<div class="p-6 rounded-2xl bg-slate-950 border border-slate-800 text-center text-sm text-slate-400">${esc(yearTag)}英雄榜查無符合「<strong class="text-white">${esc(q)}</strong>」的紀錄</div>`
        : `<div class="p-6 rounded-2xl bg-slate-950 border border-slate-800 text-center text-sm text-slate-400">${esc(yearTag)}英雄榜尚無紀錄</div>`;
      return;
    }

    const kmLabel = `${yearTag}里程`;
    listEl.innerHTML = shown.slice(0, 100).map((h) => {
      const rankBadge = h.heroRank === 1
        ? '<span class="text-2xl">🥇</span>'
        : h.heroRank === 2
          ? '<span class="text-2xl">🥈</span>'
          : h.heroRank === 3
            ? '<span class="text-2xl">🥉</span>'
            : `<span class="px-2 py-0.5 rounded-lg bg-slate-800 font-mono-num text-xs font-black text-amber-300">#${h.heroRank}</span>`;

      return `
        <details class="rounded-2xl bg-slate-950 border ${h.heroRank <= 3 ? 'border-amber-500/50' : 'border-slate-800'} p-3.5 sm:p-4" ${q ? 'open' : ''}>
          <summary class="cursor-pointer list-none flex flex-wrap items-center justify-between gap-3 select-none">
            <div class="flex items-center gap-3 min-w-0">
              <div class="w-9 text-center shrink-0">${rankBadge}</div>
              <div class="min-w-0">
                <div class="flex items-center gap-2 flex-wrap">
                  <span class="text-base sm:text-lg font-black text-white">${esc(h.name)}</span>
                  <span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/15 border border-amber-500/30 text-amber-300 font-mono-num">${esc(yearTag)}</span>
                </div>
                <div class="text-xs text-slate-400 font-mono-num mt-0.5">
                  完賽 <strong class="text-emerald-400">${h.completed}</strong> / 參賽 ${h.records.length} 場 ｜ 單場最高 ${h.bestSingleKm.toFixed(1)} km
                </div>
              </div>
            </div>
            <div class="flex items-center gap-3 ml-auto">
              <div class="text-right">
                <div class="text-[10px] text-slate-400 font-bold">${esc(kmLabel)}</div>
                <div class="text-lg sm:text-xl font-black font-mono-num text-amber-400">${h.totalKm.toFixed(2)} km</div>
              </div>
              <span class="text-xs text-sky-400 font-bold">明細 ▾</span>
            </div>
          </summary>
          <div class="mt-3 pt-3 border-t border-slate-800 space-y-2">
            ${h.records.map(renderUnifiedRecordRow).join('')}
          </div>
        </details>
      `;
    }).join('');
  }

  /* ========================================================================
   * 模式 4：🏅 公開選手證書列印（官方中英文對照樣式 ＋ 客製化上傳底圖與個人照片）
   * ======================================================================== */

  function makeCertEntryKey(e) {
    return `${e.eventKey}__${e.bib || ''}__${e.name || ''}`;
  }

  function getPrintableEntries() {
    // 證書列印僅包含「已結束」與「歷年」賽事，進行中與籌備中活動一律排除
    const all = buildUnifiedEntries({ includeLive: false, includeFinished: true, includeImported: true });
    return all
      .filter((e) => e.completed)
      .sort((a, b) =>
        String(b.eventDate || '').localeCompare(String(a.eventDate || ''))
        || (a.overallRank || 99999) - (b.overallRank || 99999)
      );
  }

  function buildCertPayloadFromEntry(entry) {
    if (!entry) return null;
    if (entry.eventId) {
      const sysEv = state.events.find((ev) => ev.eventId === entry.eventId);
      if (sysEv && entry.systemRow) {
        return { event: sysEv, row: entry.systemRow };
      }
    }
    // 歷年匯入成績轉成標準證書資料結構
    const isLap = Number(entry.laps || 0) > 0;
    const syntheticEvent = {
      eventId: entry.eventKey,
      name: entry.eventName,
      englishName: '',
      organization: 'Taiwan Ultra Runners',
      date: entry.eventDate,
      location: '',
      eventType: isLap ? 'lap_timed' : 'distance',
      lapDistanceKm: isLap && entry.laps > 0 ? Number((entry.distanceKm / entry.laps).toFixed(2)) : 0
    };
    const syntheticRow = isLap
      ? {
          bib: entry.bib || '--',
          name: entry.name,
          categoryName: entry.categoryName || '一般組',
          lapCount: entry.laps,
          totalDistanceKm: Number(entry.distanceKm || 0),
          totalElapsedMs: entry.finishTimeMs,
          overallRank: entry.overallRank,
          categoryRank: entry.categoryRank,
          avgPace: entry.avgPace || '--',
          laps: []
        }
      : {
          bib: entry.bib || '--',
          name: entry.name,
          categoryName: entry.categoryName || '一般組',
          distanceKm: Number(entry.distanceKm || 0),
          gunDurationMs: entry.finishTimeMs,
          overallRank: entry.overallRank,
          categoryRank: entry.categoryRank,
          avgPace: entry.avgPace || '--',
          status: 'finished',
          splits: []
        };
    return { event: syntheticEvent, row: syntheticRow };
  }

  async function renderPublicCertificateTab() {
    const allPrintable = getPrintableEntries();
    state.cert.entriesCache = allPrintable;

    // 整理所有可印證書之活動明細（僅列已結束與歷年賽事）
    const evMap = new Map();
    for (const e of allPrintable) {
      if (!evMap.has(e.eventKey)) {
        evMap.set(e.eventKey, {
          eventKey: e.eventKey,
          year: e.year || getEventYear(e.eventDate),
          eventDate: e.eventDate,
          eventName: e.eventName,
          sourceType: e.sourceType,
          count: 0
        });
      }
      evMap.get(e.eventKey).count += 1;
    }
    const allEvList = Array.from(evMap.values()).sort((a, b) =>
      String(b.eventDate || '').localeCompare(String(a.eventDate || ''))
    );

    // 1. 年度下拉選單
    const years = Array.from(new Set(allEvList.map((ev) => ev.year).filter(Boolean))).sort((a, b) => b.localeCompare(a));
    if (state.cert.year !== 'ALL' && !years.includes(state.cert.year)) {
      state.cert.year = 'ALL';
    }
    const yearSel = $('pub-cert-year-select');
    if (yearSel) {
      yearSel.innerHTML =
        `<option value="ALL">全部年度 (${allEvList.length} 場)</option>` +
        years.map((y) => {
          const cnt = allEvList.filter((ev) => ev.year === y).length;
          return `<option value="${esc(y)}">${esc(y)} 年度 (${cnt} 場)</option>`;
        }).join('');
      yearSel.value = state.cert.year;
    }

    // 2. 依所選年度過濾活動明細下拉選單與快速點選膠囊
    const evList = allEvList.filter((ev) => state.cert.year === 'ALL' || ev.year === state.cert.year);
    if (state.cert.eventKey !== 'ALL' && !evList.some((ev) => ev.eventKey === state.cert.eventKey)) {
      state.cert.eventKey = 'ALL';
    }

    const evSel = $('pub-cert-event-select');
    if (evSel) {
      const allLabel = state.cert.year === 'ALL'
        ? `全部已結束與歷年活動（共 ${evList.length} 場）`
        : `${state.cert.year} 年度全部活動（共 ${evList.length} 場）`;
      evSel.innerHTML =
        `<option value="ALL">${esc(allLabel)}</option>` +
        evList.map((ev) => {
          const icon = is800EventName(ev.eventName) ? '⚔️' : (ev.sourceType === 'imported' ? '📜' : '🏁');
          return `<option value="${esc(ev.eventKey)}">${icon} ${esc(ev.eventDate || '--')}｜${esc(ev.eventName)} (${ev.count}人)</option>`;
        }).join('');
      evSel.value = state.cert.eventKey;
    }

    const evPills = $('pub-cert-event-pills');
    if (evPills) {
      if (!evList.length) {
        evPills.innerHTML = '';
      } else {
        const allActive = state.cert.eventKey === 'ALL';
        evPills.innerHTML =
          `<button type="button" data-pick-cert-event="ALL" class="px-2.5 py-1 rounded-xl text-xs font-black transition ${allActive ? 'bg-sky-500 text-slate-950 shadow' : 'bg-slate-950 border border-slate-700 text-slate-300 hover:border-sky-400'}">全部 (${evList.length} 場)</button>` +
          evList.map((ev) => {
            const active = state.cert.eventKey === ev.eventKey;
            const icon = is800EventName(ev.eventName) ? '⚔️' : (ev.sourceType === 'imported' ? '📜' : '🏁');
            return `
              <button type="button" data-pick-cert-event="${esc(ev.eventKey)}"
                class="px-2.5 py-1 rounded-xl text-xs transition flex items-center gap-1.5 ${active ? 'bg-emerald-500 text-slate-950 font-black shadow' : 'bg-slate-950 border border-slate-800 hover:border-emerald-400/70 text-slate-200 font-bold'}">
                <span>${icon}</span>
                <span class="font-mono-num opacity-80">${esc(ev.eventDate || '--')}</span>
                <span class="truncate max-w-[200px]">${esc(ev.eventName)}</span>
                <span class="font-mono-num text-[11px] opacity-80">(${ev.count})</span>
              </button>
            `;
          }).join('');
      }
    }

    const scopedCount = allPrintable.filter((e) => {
      if (state.cert.year !== 'ALL' && (e.year || getEventYear(e.eventDate)) !== state.cert.year) return false;
      if (state.cert.eventKey !== 'ALL' && e.eventKey !== state.cert.eventKey) return false;
      return true;
    }).length;
    const badge = $('pub-cert-total-badge');
    if (badge) badge.textContent = `可印製 ${scopedCount} 筆完賽成績（${evList.length} 場）`;

    if ($('pub-cert-search-input')) $('pub-cert-search-input').value = state.cert.query;
    await runPublicCertSearch();
  }

  async function runPublicCertSearch() {
    const all = state.cert.entriesCache || [];
    const q = (state.cert.query || '').trim();
    const qLower = q.toLowerCase();

    const matched = all.filter((e) => {
      if (state.cert.year !== 'ALL' && (e.year || getEventYear(e.eventDate)) !== state.cert.year) return false;
      if (state.cert.eventKey !== 'ALL' && e.eventKey !== state.cert.eventKey) return false;
      if (!q) return true;
      return String(e.name || '').toLowerCase().includes(qLower)
        || String(e.eventName || '').toLowerCase().includes(qLower);
    });

    if (!matched.some((e) => makeCertEntryKey(e) === state.cert.selectedEntryKey)) {
      state.cert.selectedEntryKey = matched[0] ? makeCertEntryKey(matched[0]) : '';
    }

    const listEl = $('pub-cert-matches');
    if (listEl) {
      if (!matched.length) {
        listEl.innerHTML = '<div class="text-xs text-slate-500 py-1">查無符合條件的完賽成績，請嘗試調整關鍵字或切換賽事。</div>';
      } else {
        listEl.innerHTML = matched.slice(0, 30).map((e) => {
          const key = makeCertEntryKey(e);
          const active = key === state.cert.selectedEntryKey;
          const perf = e.laps > 0
            ? `${e.distanceKm.toFixed(1)}km/${e.laps}圈`
            : (e.finishTimeText || `${e.distanceKm}km`);
          return `
            <button type="button" data-pick-cert-entry="${esc(key)}"
              class="px-3 py-1.5 rounded-xl border text-xs transition flex items-center gap-1.5 ${active ? 'bg-emerald-500/20 border-emerald-400 text-white font-black' : 'bg-slate-950 border-slate-800 hover:border-slate-600 text-slate-300 font-bold'}">
              <span class="font-mono-num text-amber-300">${e.bib ? `#${esc(e.bib)}` : ''}</span>
              <span>${esc(e.name)}</span>
              <span class="text-[10px] text-slate-400 truncate max-w-[140px]">${esc(e.eventName)}</span>
              <span class="font-mono-num text-emerald-400">${esc(perf)}</span>
            </button>
          `;
        }).join('');
      }
    }

    const selectedEntry = matched.find((e) => makeCertEntryKey(e) === state.cert.selectedEntryKey) || null;
    $('pub-cert-empty').classList.toggle('hidden', Boolean(selectedEntry));
    $('pub-cert-studio').classList.toggle('hidden', !selectedEntry);
    if (!selectedEntry) return;

    renderPublicCertStyleUI();
    await renderPublicCertCanvas(selectedEntry);
  }

  function renderPublicCertStyleUI() {
    document.querySelectorAll('.pub-cert-style-btn').forEach((btn) => {
      const active = btn.dataset.pubCertStyle === state.cert.style;
      btn.className = active
        ? (state.cert.style === 'custom'
            ? 'pub-cert-style-btn px-3.5 py-2 rounded-lg text-xs sm:text-sm font-black bg-emerald-500 text-slate-950 transition'
            : 'pub-cert-style-btn px-3.5 py-2 rounded-lg text-xs sm:text-sm font-black bg-amber-500 text-slate-950 transition')
        : 'pub-cert-style-btn px-3.5 py-2 rounded-lg text-xs sm:text-sm font-bold text-slate-300 hover:bg-slate-800 transition';
    });
    $('pub-cert-panel-official').classList.toggle('hidden', state.cert.style !== 'official');
    $('pub-cert-panel-custom').classList.toggle('hidden', state.cert.style !== 'custom');
    $('pub-cert-drag-hint').classList.toggle('hidden', state.cert.style !== 'custom');
  }

  function getSelectedPublicCertEntry() {
    return (state.cert.entriesCache || []).find((e) => makeCertEntryKey(e) === state.cert.selectedEntryKey) || null;
  }

  async function renderPublicCertToCanvas(canvas, entry) {
    if (!Cert || !canvas || !entry) return null;
    const payload = buildCertPayloadFromEntry(entry);
    if (!payload) return null;

    if (state.cert.style === 'custom') {
      return Cert.renderPersonalizedCertificate(canvas, {
        event: payload.event,
        row: payload.row,
        bgImage: state.cert.bgImage,
        photoImage: state.cert.photoImage,
        options: {
          orientation: state.cert.cOrientation,
          cardStyle: state.cert.cStyle,
          photoShape: state.cert.cShape,
          photoPosition: state.cert.cPos,
          photoZoom: state.cert.cZoom,
          photoOffsetX: state.cert.cOffsetX,
          photoOffsetY: state.cert.cOffsetY,
          customQuote: state.cert.cQuote,
          showLogo: state.cert.cShowLogo,
          showSeal: state.cert.cShowSeal,
          showSplits: state.cert.cShowSplits
        }
      });
    }

    return Cert.renderOfficialCertificate(canvas, {
      event: payload.event,
      row: payload.row,
      options: {
        orientation: state.cert.orientation,
        format: state.cert.format,
        theme: state.cert.theme,
        showSplits: state.cert.showSplits
      }
    });
  }

  async function renderPublicCertCanvas(entryOverride = null) {
    const entry = entryOverride || getSelectedPublicCertEntry();
    const canvas = $('pub-cert-canvas');
    if (!entry || !canvas || !Cert) return;

    const res = await renderPublicCertToCanvas(canvas, entry);
    state.cert.photoBox = res?.photoBox || null;

    if ($('pub-cert-active-badge')) {
      $('pub-cert-active-badge').textContent = `${entry.eventDate || '--'}｜${entry.eventName}`;
    }
    if ($('pub-cert-preview-caption')) {
      const fmtLabel = Cert?.OFFICIAL_FORMATS?.[state.cert.format]?.label || '經典燙金典藏式';
      const thmLabel = Cert?.OFFICIAL_THEMES?.[state.cert.theme]?.label || '經典象牙金';
      const styleLabel = state.cert.style === 'custom'
        ? `客製化個人證書（${state.cert.cOrientation === 'portrait' ? 'A4 直式' : 'A4 橫式'}）`
        : `官方公版・${fmtLabel}・${thmLabel}（${state.cert.orientation === 'portrait' ? 'A4 直式' : 'A4 橫式'}）`;
      $('pub-cert-preview-caption').textContent = `${styleLabel} ｜ ${entry.name}${entry.bib ? ` (#${entry.bib})` : ''} • ${entry.categoryName || ''}`;
    }
  }

  /* ========================================================================
   * 事件綁定
   * ======================================================================== */
  function bindEvents() {
    $('btn-pub-refresh').addEventListener('click', () => loadData(false));

    document.querySelectorAll('.pub-mode-btn').forEach((btn) => {
      btn.addEventListener('click', () => switchMode(btn.dataset.pubMode));
    });

    $('btn-goto-history-from-empty').addEventListener('click', () => switchMode('history'));
    $('btn-goto-800-from-empty').addEventListener('click', () => switchMode('warriors800'));

    $('pub-event-select').addEventListener('change', (e) => {
      selectPublicEvent(e.target.value);
    });

    $('pub-event-pills').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pill-event]');
      if (!btn) return;
      selectPublicEvent(btn.dataset.pillEvent);
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

    $('pub-leaderboard-tbody').addEventListener('click', (e) => {
      const tr = e.target.closest('[data-pick-name]');
      if (!tr) return;
      const pickedName = tr.dataset.pickName || '';
      $('pub-search-input').value = pickedName;
      runSearch(pickedName);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    const handleCrossJump = (e) => {
      const certBtn = e.target.closest('[data-open-cert-eventkey]');
      if (certBtn) {
        const evKey = certBtn.dataset.openCertEventkey || 'ALL';
        const bib = certBtn.dataset.openCertBib || '';
        const name = certBtn.dataset.openCertName || '';
        const matchedEv = getPrintableEntries().find((x) => x.eventKey === evKey);
        state.cert.year = matchedEv ? (matchedEv.year || getEventYear(matchedEv.eventDate) || 'ALL') : 'ALL';
        state.cert.eventKey = evKey;
        state.cert.query = name;
        state.cert.selectedEntryKey = `${evKey}__${bib}__${name}`;
        state.mode = 'certificate';
        renderModeUI();
        renderPublicCertificateTab();
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }

      const w800Btn = e.target.closest('[data-open-800-hero]');
      if (w800Btn) {
        const name = w800Btn.dataset.open800Hero || '';
        const all800 = buildUnifiedEntries({ includeLive: true, includeFinished: true, includeImported: true, only800: true });
        const matchedYears = Array.from(
          new Set(all800.filter((r) => String(r.name || '') === name && r.year).map((r) => r.year))
        ).sort((a, b) => b.localeCompare(a));
        state.w800HeroQuery = name;
        if (matchedYears.length) {
          state.w800Year = matchedYears[0];
          state.w800HeroYear = matchedYears[0];
        } else {
          state.w800HeroYear = state.w800Year || '';
        }
        $('pub-800-hero-search').value = name;
        state.mode = 'warriors800';
        renderModeUI();
        renderWarriors800Tab();
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }

      const pickHistEvBtn = e.target.closest('[data-pick-history-event]');
      if (pickHistEvBtn) {
        state.historySubMode = 'event';
        state.historyEventKey = pickHistEvBtn.dataset.pickHistoryEvent || 'ALL';
        state.historyEventCat = 'ALL';
        state.historyEventQuery = '';
        if ($('pub-history-event-search')) $('pub-history-event-search').value = '';
        renderHistoryFilters();
        renderHistoryEventView();
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }

      const backHistEvBtn = e.target.closest('[data-back-history-events]');
      if (backHistEvBtn) {
        state.historySubMode = 'event';
        state.historyEventKey = 'ALL';
        state.historyEventCat = 'ALL';
        state.historyEventQuery = '';
        if ($('pub-history-event-search')) $('pub-history-event-search').value = '';
        renderHistoryFilters();
        renderHistoryEventView();
        return;
      }

      const histBtn = e.target.closest('[data-open-history-query]');
      if (histBtn) {
        const q = histBtn.dataset.openHistoryQuery || '';
        const idHash = histBtn.dataset.openHistoryIdhash || '';
        state.historySubMode = 'person';
        state.historyYear = 'ALL';
        state.historyEventKey = 'ALL';
        state.historyQuery = q;
        $('pub-history-input').value = q;
        state.mode = 'history';
        renderModeUI();
        renderHistoryFilters();
        runHistorySearch(q, { forceIdHash: idHash });
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
      }

      const btn = e.target.closest('[data-switch-pub-event]');
      if (!btn) return;
      selectPublicEvent(btn.dataset.switchPubEvent);
    };

    $('pub-search-results').addEventListener('click', handleCrossJump);
    $('pub-history-results').addEventListener('click', handleCrossJump);
    $('pub-800-hero-list').addEventListener('click', handleCrossJump);

    // 歷年成績查詢事件（年度切換 ＋ 個人成績查詢 ／ 場次成績查詢雙模式）
    document.querySelectorAll('.pub-history-sub-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.historySubMode = btn.dataset.historySub === 'event' ? 'event' : 'person';
        renderHistoryFilters();
        runHistorySearch($('pub-history-input').value);
      });
    });

    $('pub-history-year-filter').addEventListener('change', (e) => {
      state.historyYear = e.target.value;
      state.historyEventKey = 'ALL';
      state.historyEventCat = 'ALL';
      state.historyEventQuery = '';
      if ($('pub-history-event-search')) $('pub-history-event-search').value = '';
      renderHistoryFilters();
      runHistorySearch($('pub-history-input').value);
    });

    $('pub-history-event-filter').addEventListener('change', (e) => {
      state.historyEventKey = e.target.value;
      state.historyEventCat = 'ALL';
      state.historyEventQuery = '';
      if ($('pub-history-event-search')) $('pub-history-event-search').value = '';
      renderHistorySubModeUI();
      renderHistoryEventView();
    });

    $('pub-history-event-cat-filter')?.addEventListener('change', (e) => {
      state.historyEventCat = e.target.value;
      renderHistoryEventView();
    });

    $('pub-history-event-search')?.addEventListener('input', (e) => {
      state.historyEventQuery = e.target.value;
      renderHistoryEventView();
    });

    $('form-pub-history-search').addEventListener('submit', (e) => {
      e.preventDefault();
      state.historySubMode = 'person';
      renderHistorySubModeUI();
      runHistorySearch($('pub-history-input').value);
    });

    $('pub-history-input').addEventListener('input', (e) => {
      if (!e.target.value.trim()) runHistorySearch('');
    });

    // 全系統字體循環切換按鈕（與首頁共享設定）
    $('btn-pub-font-cycle')?.addEventListener('click', () => {
      const cur = getSavedFontScale();
      const idx = FONT_SCALE_OPTIONS.indexOf(cur);
      const next = FONT_SCALE_OPTIONS[(idx + 1) % FONT_SCALE_OPTIONS.length];
      applyPublicFontScale(next, true);
    });

    // 八百壯士系統事件
    $('pub-800-year-select').addEventListener('change', (e) => {
      state.w800Year = e.target.value;
      state.w800HeroYear = e.target.value;
      state.w800Modal.step = 'events';
      state.w800Modal.selectedEventKey = '';
      renderWarriors800Tab();
    });

    $('pub-800-hero-year-select')?.addEventListener('change', (e) => {
      state.w800HeroYear = e.target.value;
      renderWarriors800Tab();
    });

    $('pub-800-hero-year-pills')?.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-w800-hero-year]');
      if (!btn) return;
      const picked = btn.dataset.w800HeroYear || '';
      if (picked) {
        state.w800HeroYear = picked;
        renderWarriors800Tab();
      }
    });

    $('btn-pub-800-open-events')?.addEventListener('click', openWarriors800EventsModal);
    $('btn-close-modal-800')?.addEventListener('click', closeWarriors800EventsModal);
    $('modal-800-events')?.addEventListener('click', (e) => {
      if (e.target === $('modal-800-events')) closeWarriors800EventsModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && state.w800Modal.open) closeWarriors800EventsModal();
    });
    $('btn-modal-800-back')?.addEventListener('click', () => {
      state.w800Modal.step = 'events';
      state.w800Modal.selectedEventKey = '';
      state.w800Modal.runnerQuery = '';
      if ($('modal-800-runner-search')) $('modal-800-runner-search').value = '';
      renderWarriors800EventsModal();
    });
    $('modal-800-step-events')?.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-w800-event-key]');
      if (!btn) return;
      state.w800Modal.step = 'runners';
      state.w800Modal.selectedEventKey = btn.dataset.w800EventKey || '';
      state.w800Modal.runnerQuery = '';
      if ($('modal-800-runner-search')) $('modal-800-runner-search').value = '';
      renderWarriors800EventsModal();
    });
    $('modal-800-runner-search')?.addEventListener('input', (e) => {
      state.w800Modal.runnerQuery = e.target.value;
      renderWarriors800EventsModal();
    });

    $('form-pub-800-hero').addEventListener('submit', (e) => {
      e.preventDefault();
      state.w800HeroQuery = $('pub-800-hero-search').value.trim();
      renderWarriors800HeroList();
    });

    $('pub-800-hero-search').addEventListener('input', (e) => {
      state.w800HeroQuery = e.target.value.trim();
      if (!state.w800HeroQuery) renderWarriors800HeroList();
    });

    // 證書列印與客製化事件（年度選擇 ＋ 活動明細選擇 ＋ 快速點選膠囊）
    $('pub-cert-year-select')?.addEventListener('change', (e) => {
      state.cert.year = e.target.value;
      state.cert.eventKey = 'ALL';
      renderPublicCertificateTab();
    });
    $('pub-cert-event-select').addEventListener('change', (e) => {
      state.cert.eventKey = e.target.value;
      renderPublicCertificateTab();
    });
    $('pub-cert-event-pills')?.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pick-cert-event]');
      if (!btn) return;
      state.cert.eventKey = btn.dataset.pickCertEvent || 'ALL';
      renderPublicCertificateTab();
    });
    $('form-pub-cert-search').addEventListener('submit', (e) => {
      e.preventDefault();
      state.cert.query = $('pub-cert-search-input').value.trim();
      runPublicCertSearch();
    });
    $('pub-cert-search-input').addEventListener('input', (e) => {
      state.cert.query = e.target.value.trim();
      if (!state.cert.query) runPublicCertSearch();
    });
    $('pub-cert-matches').addEventListener('click', (e) => {
      const btn = e.target.closest('[data-pick-cert-entry]');
      if (!btn) return;
      state.cert.selectedEntryKey = btn.dataset.pickCertEntry;
      runPublicCertSearch();
    });

    document.querySelectorAll('.pub-cert-style-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.cert.style = btn.dataset.pubCertStyle === 'custom' ? 'custom' : 'official';
        renderPublicCertStyleUI();
        renderPublicCertCanvas();
      });
    });

    // 官方樣式控制
    $('pub-cert-orientation').addEventListener('change', (e) => { state.cert.orientation = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-format')?.addEventListener('change', (e) => { state.cert.format = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-theme').addEventListener('change', (e) => { state.cert.theme = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-show-splits').addEventListener('change', (e) => { state.cert.showSplits = e.target.checked; renderPublicCertCanvas(); });

    // 客製化：上傳底圖與上傳個人照片
    const readFileAsImage = (file) => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async () => {
        const img = await Cert.loadImage(reader.result);
        if (!img) reject(new Error('不支援的圖片格式'));
        else resolve(img);
      };
      reader.onerror = () => reject(new Error('圖片讀取失敗'));
      reader.readAsDataURL(file);
    });

    $('input-pub-cert-bg').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const img = await readFileAsImage(file);
        state.cert.bgImage = img;
        $('pub-cert-bg-status').textContent = `✅ 已套用底圖：${file.name} (${img.naturalWidth}×${img.naturalHeight})`;
        $('btn-pub-cert-clear-bg').classList.remove('hidden');
        await renderPublicCertCanvas();
      } catch (err) {
        alert(err.message);
      }
      e.target.value = '';
    });

    $('btn-pub-cert-clear-bg').addEventListener('click', () => {
      state.cert.bgImage = null;
      $('pub-cert-bg-status').textContent = '未上傳時自動套用官方經典底紋';
      $('btn-pub-cert-clear-bg').classList.add('hidden');
      renderPublicCertCanvas();
    });

    $('input-pub-cert-photo').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const img = await readFileAsImage(file);
        state.cert.photoImage = img;
        state.cert.cOffsetX = 0;
        state.cert.cOffsetY = 0;
        $('pub-cert-photo-status').textContent = `✅ 已套用照片：${file.name} (${img.naturalWidth}×${img.naturalHeight})`;
        $('btn-pub-cert-clear-photo').classList.remove('hidden');
        await renderPublicCertCanvas();
      } catch (err) {
        alert(err.message);
      }
      e.target.value = '';
    });

    $('btn-pub-cert-clear-photo').addEventListener('click', () => {
      state.cert.photoImage = null;
      state.cert.cOffsetX = 0;
      state.cert.cOffsetY = 0;
      $('pub-cert-photo-status').textContent = '支援衝線照／完賽獎牌照（可拖曳構圖）';
      $('btn-pub-cert-clear-photo').classList.add('hidden');
      renderPublicCertCanvas();
    });

    // 客製化版面與相框調整
    $('pub-cert-c-orientation').addEventListener('change', (e) => { state.cert.cOrientation = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-c-style').addEventListener('change', (e) => { state.cert.cStyle = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-c-shape').addEventListener('change', (e) => { state.cert.cShape = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-c-pos').addEventListener('change', (e) => { state.cert.cPos = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-c-zoom').addEventListener('input', (e) => {
      const pct = Number(e.target.value || 100);
      state.cert.cZoom = pct / 100;
      $('pub-cert-c-zoom-label').textContent = `${pct}%`;
      renderPublicCertCanvas();
    });
    $('btn-pub-cert-reset-photo').addEventListener('click', () => {
      state.cert.cZoom = 1;
      state.cert.cOffsetX = 0;
      state.cert.cOffsetY = 0;
      $('pub-cert-c-zoom').value = '100';
      $('pub-cert-c-zoom-label').textContent = '100%';
      renderPublicCertCanvas();
    });
    $('pub-cert-c-quote').addEventListener('input', (e) => { state.cert.cQuote = e.target.value; renderPublicCertCanvas(); });
    $('pub-cert-c-show-logo').addEventListener('change', (e) => { state.cert.cShowLogo = e.target.checked; renderPublicCertCanvas(); });
    $('pub-cert-c-show-seal').addEventListener('change', (e) => { state.cert.cShowSeal = e.target.checked; renderPublicCertCanvas(); });
    $('pub-cert-c-show-splits').addEventListener('change', (e) => { state.cert.cShowSplits = e.target.checked; renderPublicCertCanvas(); });

    // 客製化模式：於相框內直接拖曳調整個人照片構圖
    let photoDrag = null;
    const pubCanvas = $('pub-cert-canvas');
    const getPubCanvasPt = (clientX, clientY) => {
      const rect = pubCanvas.getBoundingClientRect();
      return {
        x: ((clientX - rect.left) / rect.width) * pubCanvas.width,
        y: ((clientY - rect.top) / rect.height) * pubCanvas.height
      };
    };
    const startPhotoDrag = (clientX, clientY) => {
      if (state.cert.style !== 'custom' || !state.cert.photoImage || !state.cert.photoBox) return;
      const pt = getPubCanvasPt(clientX, clientY);
      const b = state.cert.photoBox;
      if (pt.x >= b.x && pt.x <= b.x + b.w && pt.y >= b.y && pt.y <= b.y + b.h) {
        photoDrag = { startX: pt.x, startY: pt.y, origX: state.cert.cOffsetX, origY: state.cert.cOffsetY, boxW: b.w, boxH: b.h };
      }
    };
    const movePhotoDrag = (clientX, clientY) => {
      if (!photoDrag) return;
      const pt = getPubCanvasPt(clientX, clientY);
      const dx = ((pt.x - photoDrag.startX) / photoDrag.boxW) * 100;
      const dy = ((pt.y - photoDrag.startY) / photoDrag.boxH) * 100;
      state.cert.cOffsetX = Math.max(-50, Math.min(50, photoDrag.origX + dx));
      state.cert.cOffsetY = Math.max(-50, Math.min(50, photoDrag.origY + dy));
      renderPublicCertCanvas();
    };
    const endPhotoDrag = () => { photoDrag = null; };

    pubCanvas.addEventListener('mousedown', (e) => startPhotoDrag(e.clientX, e.clientY));
    window.addEventListener('mousemove', (e) => movePhotoDrag(e.clientX, e.clientY));
    window.addEventListener('mouseup', endPhotoDrag);
    pubCanvas.addEventListener('touchstart', (e) => {
      if (e.touches[0]) {
        startPhotoDrag(e.touches[0].clientX, e.touches[0].clientY);
        if (photoDrag) e.preventDefault();
      }
    }, { passive: false });
    window.addEventListener('touchmove', (e) => {
      if (photoDrag && e.touches[0]) {
        movePhotoDrag(e.touches[0].clientX, e.touches[0].clientY);
        e.preventDefault();
      }
    }, { passive: false });
    window.addEventListener('touchend', endPhotoDrag);

    // 下載圖檔 / PDF / 直接列印
    const exportPublicCert = async (type) => {
      const entry = getSelectedPublicCertEntry();
      if (!entry) return;
      const offscreen = document.createElement('canvas');
      await renderPublicCertToCanvas(offscreen, entry);
      const tag = state.cert.style === 'custom' ? '客製化證書' : '官方完賽證書';
      const safeName = `${tag}_${entry.eventName}_${entry.bib || ''}_${entry.name}`.replace(/[\\/:*?"<>|]/g, '_');
      if (type === 'png') {
        await Cert.downloadPng(offscreen, `${safeName}.png`);
      } else if (type === 'pdf') {
        Cert.downloadPdfFromCanvas(offscreen, `${safeName}.pdf`);
      } else if (type === 'print') {
        Cert.printCanvasCertificate(offscreen, `${entry.name} - ${entry.eventName}`);
      }
    };

    $('btn-pub-cert-png').addEventListener('click', () => exportPublicCert('png').catch((err) => alert(err.message)));
    $('btn-pub-cert-pdf').addEventListener('click', () => exportPublicCert('pdf').catch((err) => alert(err.message)));
    $('btn-pub-cert-print').addEventListener('click', () => exportPublicCert('print').catch((err) => alert(err.message)));
  }

  document.addEventListener('DOMContentLoaded', initPublicPage);
})();
