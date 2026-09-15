/**
 * OmniConverter - Frontend Logic & In-Browser Multi-Format Conversion Engine
 * Runs 100% offline or on static hosts (GitHub Pages) with full client-side conversion.
 */

// Application State
const AppState = {
  rawDataset: null,
  activeCalls: [],
  filteredCalls: [],
  selectedIds: new Set(), // Starts empty: manual selection mode
  currentPage: 1,
  pageSize: 15,
  currentSort: { column: 'date', order: 'desc' },
  filterQuery: '',
  filterType: 'all',
  filterDateFrom: '',
  filterDateTo: '',
  timelineMode: 'duration', // 'duration' or 'count'
  backendAvailable: false,
  charts: {
    timeline: null,
    hourly: null
  }
};

// DOM Selectors
const DOM = {
  connectionBadge: document.getElementById('connection-badge'),
  connectionStatusText: document.getElementById('connection-status-text'),
  btnLoadSample: document.getElementById('btn-load-sample'),
  dropZone: document.getElementById('drop-zone'),
  fileInput: document.getElementById('file-input'),
  loadedFileBanner: document.getElementById('loaded-file-banner'),
  loadedFilenameDisplay: document.getElementById('loaded-filename-display'),
  loadedDetailsDisplay: document.getElementById('loaded-details-display'),
  btnReselectFile: document.getElementById('btn-reselect-file'),
  eventTitleInput: document.getElementById('event-title-input'),
  chkExportFiltered: document.getElementById('chk-export-filtered'),
  filteredCountBadge: document.getElementById('filtered-count-badge'),

  // Export buttons
  btnExportXlsx: document.getElementById('btn-export-xlsx'),
  btnExportCsv: document.getElementById('btn-export-csv'),
  btnExportIcs: document.getElementById('btn-export-ics'),
  btnExportJson: document.getElementById('btn-export-json'),
  btnExportZip: document.getElementById('btn-export-zip'),

  // KPIs
  kpiTotalCalls: document.getElementById('kpi-total-calls'),
  kpiPeriodSubtext: document.getElementById('kpi-period-subtext'),
  kpiTotalDuration: document.getElementById('kpi-total-duration'),
  kpiTotalSec: document.getElementById('kpi-total-sec'),
  kpiAvgDuration: document.getElementById('kpi-avg-duration'),
  kpiTypesDisplay: document.getElementById('kpi-types-display'),
  kpiVoiceBar: document.getElementById('kpi-voice-bar'),
  kpiAnsweredRate: document.getElementById('kpi-answered-rate'),
  kpiStatusSubtext: document.getElementById('kpi-status-subtext'),

  // Charts
  chartTimeline: document.getElementById('chart-timeline'),
  chartHourly: document.getElementById('chart-hourly'),

  // Filters & Table
  searchInput: document.getElementById('search-input'),
  btnClearSearch: document.getElementById('btn-clear-search'),
  filterDateFrom: document.getElementById('filter-date-from'),
  filterDateTo: document.getElementById('filter-date-to'),
  visibleCount: document.getElementById('visible-count'),
  totalCount: document.getElementById('total-count'),
  btnSelectAll: document.getElementById('btn-select-all'),
  btnSelectNone: document.getElementById('btn-select-none'),
  selectPageSize: document.getElementById('select-page-size'),
  callsTableBody: document.getElementById('calls-table-body'),
  chkSelectPage: document.getElementById('chk-select-page'),

  // Pagination
  paginationInfo: document.getElementById('pagination-info'),
  btnPageFirst: document.getElementById('btn-page-first'),
  btnPagePrev: document.getElementById('btn-page-prev'),
  paginationNumbers: document.getElementById('pagination-numbers'),
  btnPageNext: document.getElementById('btn-page-next'),
  btnPageLast: document.getElementById('btn-page-last'),

  // Modal & Toast
  modalDetails: document.getElementById('modal-details'),
  modalDetailsContent: document.getElementById('modal-details-content'),
  btnCloseModal: document.getElementById('btn-close-modal'),
  btnModalDone: document.getElementById('btn-modal-done'),
  btnCopyJson: document.getElementById('btn-copy-json'),
  toastContainer: document.getElementById('toast-container'),

  // Navigation Tabs
  tabBtnConverter: document.getElementById('tab-btn-converter'),
  tabBtnSlicer: document.getElementById('tab-btn-slicer'),
  viewConverter: document.getElementById('view-converter'),
  viewSlicer: document.getElementById('view-slicer'),

  // ICS Calendar Slicer
  slicerDropZone: document.getElementById('slicer-drop-zone'),
  slicerFileInput: document.getElementById('slicer-file-input'),
  slicerFileBanner: document.getElementById('slicer-file-banner'),
  slicerFilenameDisplay: document.getElementById('slicer-filename-display'),
  slicerDetailsDisplay: document.getElementById('slicer-details-display'),
  btnSlicerReselect: document.getElementById('btn-slicer-reselect'),
  slicerWorkspace: document.getElementById('slicer-workspace'),
  btnModeSize: document.getElementById('btn-mode-size'),
  btnModeEvents: document.getElementById('btn-mode-events'),
  blockSizeOptions: document.getElementById('block-size-options'),
  blockEventsOptions: document.getElementById('block-events-options'),
  sizePresetContainer: document.getElementById('size-preset-container'),
  eventsPresetContainer: document.getElementById('events-preset-container'),
  customSizeVal: document.getElementById('custom-size-val'),
  customSizeUnit: document.getElementById('custom-size-unit'),
  customEventsVal: document.getElementById('custom-events-val'),
  slicerBaseName: document.getElementById('slicer-base-name'),
  btnRecalculateSplit: document.getElementById('btn-recalculate-split'),
  slicerPlanSummary: document.getElementById('slicer-plan-summary'),
  slicerStatEvents: document.getElementById('slicer-stat-events'),
  slicerStatSize: document.getElementById('slicer-stat-size'),
  slicerStatParts: document.getElementById('slicer-stat-parts'),
  slicerStatTz: document.getElementById('slicer-stat-tz'),
  btnDownloadAllZip: document.getElementById('btn-download-all-zip'),
  slicerPartsContainer: document.getElementById('slicer-parts-container'),
  slicerPartsCount: document.getElementById('slicer-parts-count'),
  slicerPartsList: document.getElementById('slicer-parts-list')
};

// Slicer Application State
const SlicerState = {
  rawContent: '',
  filename: 'calendar.ics',
  baseName: 'calendar_part',
  sizeBytes: 0,
  headers: [],
  timezones: [],
  events: [],
  splitMode: 'size', // 'size' or 'events'
  targetSizeKB: 950,
  targetEvents: 500,
  parts: []
};

// Utilities
function showToast(message, type = 'info', duration = 3200) {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const icon = type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️';
  toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
  DOM.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function formatDuration(sec) {
  if (isNaN(sec) || sec === null) return '00:00';
  sec = Math.round(sec);
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  if (h > 0) {
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  }
  return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

function formatDurationHuman(sec) {
  if (isNaN(sec) || sec === null) return '0s';
  sec = Math.round(sec);
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  if (h > 0) {
    return `${h}h ${m}m ${s}s`;
  }
  if (m > 0) {
    return `${m}m ${s}s`;
  }
  return `${s}s`;
}

// Data Normalization
function normalizeDataset(raw) {
  let callsRaw = [];
  let contact = { name: 'Sample Contact', phone: '', jid: '' };
  let exportedAt = new Date().toISOString();

  if (Array.isArray(raw)) {
    callsRaw = raw;
  } else if (raw && typeof raw === 'object') {
    callsRaw = raw.calls || [];
    if (raw.contact) contact = raw.contact;
    if (raw.exportedAt) exportedAt = raw.exportedAt;
  }

  const normalized = callsRaw.map((c, idx) => {
    const durationSec = parseInt(c.durationSec || c.duration || 0, 10);
    const durFormatted = c.durationFormatted || formatDuration(durationSec);

    let iso = c.iso || '';
    let dateStr = c.date || '';
    let timeStr = c.time || '';

    if (!iso && dateStr && timeStr) {
      iso = `${dateStr}T${timeStr}.000Z`;
    } else if (iso && (!dateStr || !timeStr)) {
      try {
        const dt = new Date(iso);
        dateStr = dt.toISOString().split('T')[0];
        timeStr = dt.toISOString().split('T')[1].slice(0, 8);
      } catch (e) {}
    }

    let mediaType = c.mediaType || 'Voice';
    if (mediaType.toLowerCase().includes('voz') || mediaType.toLowerCase().includes('voice')) {
      mediaType = 'Voice';
    } else if (mediaType.toLowerCase().includes('vídeo') || mediaType.toLowerCase().includes('video')) {
      mediaType = 'Video';
    }

    const fromMe = Boolean(c.fromMe);
    const direction = c.direction || (fromMe ? 'Outgoing' : 'Incoming');
    let status = c.status || (fromMe ? 'Outgoing (Answered)' : 'Incoming (Answered)');

    let isMissed = Boolean(c.isMissed);
    const statusLower = status.toLowerCase();
    if (statusLower.includes('missed') || statusLower.includes('perdida') || statusLower.includes('unanswered') || statusLower.includes('não atendida')) {
      isMissed = true;
    }

    return {
      id: c.id || `call_${idx}_${Date.now()}`,
      chatId: c.chatId || '',
      timestamp: c.timestamp || 0,
      timestampMs: c.timestampMs || (c.timestamp ? c.timestamp * 1000 : 0),
      date: dateStr,
      time: timeStr,
      iso: iso,
      mediaType: mediaType,
      direction: direction,
      status: status,
      isMissed: isMissed,
      durationSec: durationSec,
      durationFormatted: durFormatted,
      fromMe: fromMe,
      source: c.source || 'whatsapp-export',
      _raw: c
    };
  });

  return {
    contact,
    exportedAt,
    calls: normalized
  };
}

// Load Data into UI
function loadDataset(dataset, filename = 'call_logs.json') {
  AppState.rawDataset = dataset;
  AppState.activeCalls = dataset.calls;
  AppState.selectedIds = new Set(); // Default: empty so user can check just 1, 2, or 3!
  AppState.currentPage = 1;

  DOM.loadedFilenameDisplay.textContent = filename;
  const totalSec = dataset.calls.reduce((sum, c) => sum + c.durationSec, 0);
  DOM.loadedDetailsDisplay.textContent = `${dataset.calls.length} calls loaded • Total duration: ${formatDurationHuman(totalSec)}`;
  DOM.loadedFileBanner.classList.remove('hidden');

  applyFilters();
  updateKPIs();
  updateCharts();
  updateSelectionUI();
  showToast(`Successfully loaded ${dataset.calls.length} calls!`, 'success');
}

// Check Backend Connection or Fallback to Static Sample Data
async function checkBackend() {
  try {
    const res = await fetch('/api/data');
    if (res.ok) {
      AppState.backendAvailable = true;
      DOM.connectionBadge.className = 'connection-pill online';
      DOM.connectionStatusText.textContent = 'Local Server Active';
      const data = await res.json();
      loadDataset(normalizeDataset(data), 'data.json (server)');
      return;
    }
  } catch (err) {}

  AppState.backendAvailable = false;
  DOM.connectionBadge.className = 'connection-pill offline';
  DOM.connectionStatusText.textContent = 'Browser Mode (Offline)';

  const samplePaths = ['data.json', 'sample_data.json', './data.json', './sample_data.json', '../data.json'];
  for (const path of samplePaths) {
    try {
      const res = await fetch(path);
      if (res.ok) {
        const json = await res.json();
        loadDataset(normalizeDataset(json), 'data.json');
        return;
      }
    } catch (e) {}
  }

  DOM.loadedFilenameDisplay.textContent = 'Ready for data';
  DOM.loadedDetailsDisplay.textContent = 'Drag & drop any WhatsApp JSON export above or click Load Sample';
  DOM.loadedFileBanner.classList.remove('hidden');
}

// KPIs Calculation
function updateKPIs() {
  const calls = AppState.filteredCalls;
  const total = calls.length;

  if (total === 0) {
    DOM.kpiTotalCalls.textContent = '0';
    DOM.kpiTotalDuration.textContent = '0s';
    DOM.kpiTotalSec.textContent = '0 seconds';
    DOM.kpiAvgDuration.textContent = '0s';
    DOM.kpiTypesDisplay.textContent = '0 Voice / 0 Video';
    DOM.kpiAnsweredRate.textContent = '0%';
    DOM.kpiStatusSubtext.textContent = 'No matching calls';
    return;
  }

  const totalSec = calls.reduce((acc, c) => acc + c.durationSec, 0);
  const avgSec = Math.round(totalSec / total);

  const voiceCount = calls.filter(c => c.mediaType.toLowerCase() === 'voice').length;
  const videoCount = calls.filter(c => c.mediaType.toLowerCase() === 'video').length;
  const missedCount = calls.filter(c => c.isMissed).length;
  const answeredCount = total - missedCount;
  const answeredPct = Math.round((answeredCount / total) * 100);

  const dates = calls.map(c => c.date).filter(Boolean).sort();
  const minDate = dates[0] || '';
  const maxDate = dates[dates.length - 1] || '';

  DOM.kpiTotalCalls.textContent = total.toLocaleString();
  DOM.kpiPeriodSubtext.textContent = minDate && maxDate ? `${minDate} to ${maxDate}` : 'N/A';
  DOM.kpiTotalDuration.textContent = formatDurationHuman(totalSec);
  DOM.kpiTotalSec.textContent = `${totalSec.toLocaleString()} seconds`;
  DOM.kpiAvgDuration.textContent = formatDurationHuman(avgSec);
  DOM.kpiTypesDisplay.textContent = `${voiceCount} Voice / ${videoCount} Video`;

  const voiceRate = total > 0 ? (voiceCount / total) * 100 : 100;
  DOM.kpiVoiceBar.style.width = `${voiceRate}%`;

  DOM.kpiAnsweredRate.textContent = `${answeredPct}%`;
  DOM.kpiStatusSubtext.textContent = `${answeredCount} Answered • ${missedCount} Missed`;
}

// Charts Rendering
function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const calls = AppState.filteredCalls;

  const dateMap = {};
  calls.forEach(c => {
    if (!c.date) return;
    if (!dateMap[c.date]) {
      dateMap[c.date] = { count: 0, durationSec: 0 };
    }
    dateMap[c.date].count += 1;
    dateMap[c.date].durationSec += c.durationSec;
  });

  const sortedDates = Object.keys(dateMap).sort();
  const timelineLabels = sortedDates.map(d => {
    const parts = d.split('-');
    return `${parts[1]}/${parts[2]}`;
  });

  const timelineData = sortedDates.map(d => {
    if (AppState.timelineMode === 'duration') {
      return +(dateMap[d].durationSec / 3600).toFixed(2);
    }
    return dateMap[d].count;
  });

  if (AppState.charts.timeline) {
    AppState.charts.timeline.destroy();
  }

  const ctxTimeline = DOM.chartTimeline.getContext('2d');
  const gradientTimeline = ctxTimeline.createLinearGradient(0, 0, 0, 240);
  gradientTimeline.addColorStop(0, 'rgba(99, 102, 241, 0.45)');
  gradientTimeline.addColorStop(1, 'rgba(99, 102, 241, 0.02)');

  AppState.charts.timeline = new Chart(ctxTimeline, {
    type: 'line',
    data: {
      labels: timelineLabels,
      datasets: [{
        label: AppState.timelineMode === 'duration' ? 'Duration (Hours)' : 'Total Calls',
        data: timelineData,
        borderColor: '#818cf8',
        backgroundColor: gradientTimeline,
        borderWidth: 2.4,
        fill: true,
        tension: 0.35,
        pointBackgroundColor: '#6366f1',
        pointRadius: timelineLabels.length > 40 ? 1 : 3,
        pointHoverRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          titleFont: { family: 'Plus Jakarta Sans', size: 12 },
          bodyFont: { family: 'Plus Jakarta Sans', size: 12 },
          padding: 10,
          borderColor: 'rgba(255, 255, 255, 0.1)',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => AppState.timelineMode === 'duration' ? `${ctx.parsed.y} hours` : `${ctx.parsed.y} calls`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.04)' },
          ticks: { color: '#64748b', maxTicksLimit: 12, font: { size: 10 } }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#64748b', font: { size: 10 } }
        }
      }
    }
  });

  const hourCounts = new Array(24).fill(0);
  calls.forEach(c => {
    if (c.time) {
      const h = parseInt(c.time.split(':')[0], 10);
      if (!isNaN(h) && h >= 0 && h < 24) {
        hourCounts[h] += 1;
      }
    }
  });

  const hourLabels = Array.from({ length: 24 }, (_, i) => `${i.toString().padStart(2, '0')}:00`);

  if (AppState.charts.hourly) {
    AppState.charts.hourly.destroy();
  }

  const ctxHourly = DOM.chartHourly.getContext('2d');
  AppState.charts.hourly = new Chart(ctxHourly, {
    type: 'bar',
    data: {
      labels: hourLabels,
      datasets: [{
        label: 'Calls by Hour',
        data: hourCounts,
        backgroundColor: 'rgba(6, 182, 212, 0.65)',
        hoverBackgroundColor: '#06b6d4',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          titleFont: { family: 'Plus Jakarta Sans', size: 12 },
          bodyFont: { family: 'Plus Jakarta Sans', size: 12 },
          padding: 8,
          borderColor: 'rgba(255, 255, 255, 0.1)',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: '#64748b', font: { size: 9 }, maxTicksLimit: 12 }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#64748b', font: { size: 9 }, precision: 0 }
        }
      }
    }
  });
}

// Update Selection UI & Export Indicators
function updateSelectionUI() {
  const selectedCount = AppState.selectedIds.size;
  const filteredCount = AppState.filteredCalls.length;

  if (selectedCount > 0) {
    DOM.filteredCountBadge.textContent = `${selectedCount} selected`;
    DOM.chkExportFiltered.closest('.checkbox-container').querySelector('.checkbox-label').innerHTML =
      `⭐ <strong>Exporting ONLY ${selectedCount} selected call(s)</strong> (Uncheck to export all ${filteredCount})`;
  } else {
    DOM.filteredCountBadge.textContent = filteredCount;
    DOM.chkExportFiltered.closest('.checkbox-container').querySelector('.checkbox-label').innerHTML =
      `Export only currently filtered calls (<strong>${filteredCount}</strong> items)`;
  }
}

// Filtering & Searching Logic
function applyFilters() {
  const query = AppState.filterQuery.toLowerCase().trim();
  const typeFilter = AppState.filterType;
  const dateFrom = AppState.filterDateFrom;
  const dateTo = AppState.filterDateTo;

  AppState.filteredCalls = AppState.activeCalls.filter(c => {
    if (query) {
      const matchId = c.id.toLowerCase().includes(query);
      const matchDate = c.date.toLowerCase().includes(query);
      const matchTime = c.time.toLowerCase().includes(query);
      const matchDur = c.durationFormatted.toLowerCase().includes(query);
      const matchType = c.mediaType.toLowerCase().includes(query);
      const matchStatus = c.status.toLowerCase().includes(query);
      if (!matchId && !matchDate && !matchTime && !matchDur && !matchType && !matchStatus) {
        return false;
      }
    }

    if (typeFilter === 'voice' && c.mediaType.toLowerCase() !== 'voice') return false;
    if (typeFilter === 'video' && c.mediaType.toLowerCase() !== 'video') return false;
    if (typeFilter === 'answered' && c.isMissed) return false;
    if (typeFilter === 'missed' && !c.isMissed) return false;
    if (typeFilter === 'long' && c.durationSec < 1800) return false;

    if (dateFrom && c.date && c.date < dateFrom) return false;
    if (dateTo && c.date && c.date > dateTo) return false;

    return true;
  });

  sortCalls();
  AppState.currentPage = 1;

  renderTable();
  updateKPIs();
  updateCharts();
  updateSelectionUI();

  DOM.visibleCount.textContent = AppState.filteredCalls.length;
  DOM.totalCount.textContent = AppState.activeCalls.length;
}

function sortCalls() {
  const col = AppState.currentSort.column;
  const order = AppState.currentSort.order === 'asc' ? 1 : -1;

  AppState.filteredCalls.sort((a, b) => {
    let valA = a[col];
    let valB = b[col];

    if (col === 'date') {
      const timeA = `${a.date}T${a.time}`;
      const timeB = `${b.date}T${b.time}`;
      return timeA.localeCompare(timeB) * order;
    }

    if (typeof valA === 'string') {
      return valA.localeCompare(valB) * order;
    }
    return (valA - valB) * order;
  });
}

// Table Rendering
function renderTable() {
  const tbody = DOM.callsTableBody;
  tbody.innerHTML = '';

  const total = AppState.filteredCalls.length;
  if (total === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; padding: 2.5rem; color: #64748b;">No call records match your filter criteria.</td></tr>`;
    DOM.paginationInfo.textContent = '0 records';
    DOM.paginationNumbers.innerHTML = '';
    return;
  }

  const pageSize = AppState.pageSize;
  const totalPages = Math.ceil(total / pageSize);
  if (AppState.currentPage > totalPages) AppState.currentPage = totalPages;
  if (AppState.currentPage < 1) AppState.currentPage = 1;

  const startIdx = (AppState.currentPage - 1) * pageSize;
  const endIdx = Math.min(startIdx + pageSize, total);
  const pageCalls = AppState.filteredCalls.slice(startIdx, endIdx);

  const allPageSelected = pageCalls.every(c => AppState.selectedIds.has(c.id));
  DOM.chkSelectPage.checked = allPageSelected && pageCalls.length > 0;

  pageCalls.forEach((call, idx) => {
    const tr = document.createElement('tr');
    const isSelected = AppState.selectedIds.has(call.id);
    if (isSelected) tr.classList.add('row-selected');

    const globalIdx = startIdx + idx + 1;
    const isVoice = call.mediaType.toLowerCase() === 'voice';
    const mediaBadgeClass = isVoice ? 'voice' : 'video';
    const mediaIcon = isVoice ? '🎙️' : '📹';

    const statusBadgeClass = call.isMissed ? 'missed' : 'answered';
    const statusIcon = call.isMissed ? '❌' : '✅';

    tr.innerHTML = `
      <td class="col-chk">
        <input type="checkbox" class="row-checkbox" data-id="${call.id}" ${isSelected ? 'checked' : ''}>
      </td>
      <td class="col-num font-mono text-muted">${globalIdx}</td>
      <td class="col-date font-mono">${call.date}</td>
      <td class="col-time font-mono">${call.time}</td>
      <td class="col-type">
        <span class="badge-media ${mediaBadgeClass}">${mediaIcon} ${call.mediaType}</span>
      </td>
      <td class="col-dir">${call.direction}</td>
      <td class="col-status">
        <span class="badge-status ${statusBadgeClass}">${statusIcon} ${call.status}</span>
      </td>
      <td class="col-dur">
        <span class="duration-pill">${call.durationFormatted}</span>
      </td>
      <td class="col-iso font-mono text-subtle" style="font-size: 0.76rem;">${call.iso.slice(0, 19)}Z</td>
      <td class="col-actions">
        <button class="btn-icon-action btn-inspect-row" data-id="${call.id}" title="View details">🔍</button>
        <button class="btn-icon-action btn-export-single-ics" data-id="${call.id}" title="Export only this event to .ICS">📅</button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  DOM.paginationInfo.textContent = `Showing ${startIdx + 1} to ${endIdx} of ${total} calls (Page ${AppState.currentPage} of ${totalPages})`;
  renderPaginationButtons(totalPages);
}

function renderPaginationButtons(totalPages) {
  const container = DOM.paginationNumbers;
  container.innerHTML = '';

  DOM.btnPageFirst.disabled = AppState.currentPage <= 1;
  DOM.btnPagePrev.disabled = AppState.currentPage <= 1;
  DOM.btnPageNext.disabled = AppState.currentPage >= totalPages;
  DOM.btnPageLast.disabled = AppState.currentPage >= totalPages;

  let startPage = Math.max(1, AppState.currentPage - 2);
  let endPage = Math.min(totalPages, startPage + 4);
  if (endPage - startPage < 4) {
    startPage = Math.max(1, endPage - 4);
  }

  for (let p = startPage; p <= endPage; p++) {
    const btn = document.createElement('button');
    btn.className = `btn-page ${p === AppState.currentPage ? 'active' : ''}`;
    btn.textContent = p;
    btn.addEventListener('click', () => {
      AppState.currentPage = p;
      renderTable();
    });
    container.appendChild(btn);
  }
}

// Modal View
function openDetailsModal(callId) {
  const call = AppState.activeCalls.find(c => c.id === callId);
  if (!call) return;

  const content = DOM.modalDetailsContent;
  content.innerHTML = `
    <div class="modal-detail-row">
      <span class="modal-detail-label">Record ID:</span>
      <span class="modal-detail-val font-mono" style="font-size: 0.8rem; word-break: break-all;">${call.id}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Date & Time:</span>
      <span class="modal-detail-val">${call.date} at ${call.time}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">UTC Start (ISO):</span>
      <span class="modal-detail-val font-mono">${call.iso}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Duration:</span>
      <span class="modal-detail-val text-amber font-mono">${call.durationFormatted} (${call.durationSec}s)</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Media Type:</span>
      <span class="modal-detail-val">${call.mediaType}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Direction:</span>
      <span class="modal-detail-val">${call.direction}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Status:</span>
      <span class="modal-detail-val">${call.status}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Chat ID:</span>
      <span class="modal-detail-val font-mono" style="font-size: 0.8rem;">${call.chatId}</span>
    </div>
    <div class="modal-detail-row">
      <span class="modal-detail-label">Source:</span>
      <span class="modal-detail-val text-muted">${call.source}</span>
    </div>
  `;

  DOM.btnCopyJson.onclick = () => {
    navigator.clipboard.writeText(JSON.stringify(call._raw || call, null, 2));
    showToast('Record JSON copied to clipboard!', 'success');
  };

  DOM.modalDetails.classList.remove('hidden');
}

function closeModal() {
  DOM.modalDetails.classList.add('hidden');
}

// Multi-Format Exporters (Client-Side Standalone & API)
function getTargetCallsForExport() {
  // If user selected specific calls with checkboxes, return ONLY those calls!
  if (AppState.selectedIds.size > 0) {
    const selected = AppState.activeCalls.filter(c => AppState.selectedIds.has(c.id));
    if (selected.length > 0) return selected;
  }

  const onlyFiltered = DOM.chkExportFiltered.checked;
  return onlyFiltered ? AppState.filteredCalls : AppState.activeCalls;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

// 1. Excel (.xlsx) Exporter
function exportToExcel() {
  const calls = getTargetCallsForExport();
  if (calls.length === 0) {
    showToast('No calls available to export!', 'error');
    return;
  }

  const rows = calls.map((c, i) => ({
    "No.": i + 1,
    "Date": c.date,
    "Time": c.time,
    "Type": c.mediaType,
    "Direction": c.direction,
    "Status": c.status,
    "Duration (seconds)": c.durationSec,
    "Formatted Duration": c.durationFormatted,
    "UTC Start (ISO)": c.iso,
    "Record ID": c.id,
    "Chat ID": c.chatId
  }));

  const wb = XLSX.utils.book_new();
  const wsCalls = XLSX.utils.json_to_sheet(rows);

  wsCalls['!cols'] = [
    { wch: 6 },  // No.
    { wch: 12 }, // Date
    { wch: 10 }, // Time
    { wch: 10 }, // Type
    { wch: 12 }, // Direction
    { wch: 24 }, // Status
    { wch: 18 }, // Duration (s)
    { wch: 18 }, // Formatted Duration
    { wch: 26 }, // ISO
    { wch: 35 }, // ID
    { wch: 30 }  // Chat ID
  ];
  XLSX.utils.book_append_sheet(wb, wsCalls, "Call Log");

  const totalSec = calls.reduce((sum, c) => sum + c.durationSec, 0);
  const voiceCount = calls.filter(c => c.mediaType.toLowerCase() === 'voice').length;
  const videoCount = calls.filter(c => c.mediaType.toLowerCase() === 'video').length;
  const missedCount = calls.filter(c => c.isMissed).length;

  const statsRows = [
    { "Metric": "Total Exported Calls", "Value": calls.length },
    { "Metric": "Total Duration (seconds)", "Value": totalSec },
    { "Metric": "Total Duration Formatted", "Value": formatDuration(totalSec) },
    { "Metric": "Average Duration per Call", "Value": formatDuration(Math.round(totalSec / calls.length)) },
    { "Metric": "Voice Calls", "Value": voiceCount },
    { "Metric": "Video Calls", "Value": videoCount },
    { "Metric": "Answered Calls", "Value": calls.length - missedCount },
    { "Metric": "Missed Calls", "Value": missedCount },
    { "Metric": "Export Timestamp", "Value": new Date().toISOString() }
  ];

  const wsStats = XLSX.utils.json_to_sheet(statsRows);
  wsStats['!cols'] = [{ wch: 30 }, { wch: 26 }];
  XLSX.utils.book_append_sheet(wb, wsStats, "Statistical Summary");

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fname = calls.length === AppState.activeCalls.length ? 'call_logs.xlsx' : `call_logs_${calls.length}_selected.xlsx`;
  downloadBlob(blob, fname);
  showToast(`Excel (.xlsx) generated successfully (${calls.length} calls)!`, 'success');
}

// 2. CSV (.csv) Exporter with UTF-8 BOM
function exportToCsv() {
  const calls = getTargetCallsForExport();
  if (calls.length === 0) {
    showToast('No calls available to export!', 'error');
    return;
  }

  const delimiter = ';';
  const headers = [
    "ID", "Date", "Time", "ISO_UTC", "Media_Type", "Direction",
    "Status", "Missed", "Duration_Seconds", "Duration_Formatted",
    "Chat_ID", "From_Me"
  ];

  const csvRows = [headers.join(delimiter)];

  calls.forEach(c => {
    const row = [
      `"${c.id}"`,
      `"${c.date}"`,
      `"${c.time}"`,
      `"${c.iso}"`,
      `"${c.mediaType}"`,
      `"${c.direction}"`,
      `"${c.status}"`,
      `"${c.isMissed ? 'Yes' : 'No'}"`,
      c.durationSec,
      `"${c.durationFormatted}"`,
      `"${c.chatId}"`,
      `"${c.fromMe ? 'Yes' : 'No'}"`
    ];
    csvRows.push(row.join(delimiter));
  });

  const csvContent = '\uFEFF' + csvRows.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const fname = calls.length === AppState.activeCalls.length ? 'call_logs.csv' : `call_logs_${calls.length}_selected.csv`;
  downloadBlob(blob, fname);
  showToast(`CSV (.csv) generated successfully (${calls.length} calls)!`, 'success');
}

// 3. JSON (.json) Exporter
function exportToJson() {
  const calls = getTargetCallsForExport();
  if (calls.length === 0) {
    showToast('No calls available to export!', 'error');
    return;
  }

  const totalSec = calls.reduce((s, c) => s + c.durationSec, 0);
  const payload = {
    exportedAt: new Date().toISOString(),
    contact: AppState.rawDataset ? AppState.rawDataset.contact : {},
    totalCalls: calls.length,
    totalDurationSeconds: totalSec,
    totalDurationFormatted: formatDuration(totalSec),
    calls: calls.map(c => ({
      id: c.id,
      chatId: c.chatId,
      timestamp: c.timestamp,
      timestampMs: c.timestampMs,
      date: c.date,
      time: c.time,
      iso: c.iso,
      mediaType: c.mediaType,
      direction: c.direction,
      status: c.status,
      isMissed: c.isMissed,
      durationSec: c.durationSec,
      durationFormatted: c.durationFormatted,
      fromMe: c.fromMe,
      source: c.source
    }))
  };

  const jsonStr = JSON.stringify(payload, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
  const fname = calls.length === AppState.activeCalls.length ? 'call_logs.json' : `call_logs_${calls.length}_selected.json`;
  downloadBlob(blob, fname);
  showToast(`JSON (.json) generated successfully (${calls.length} calls)!`, 'success');
}

// 4. iCalendar (.ics) Exporter
function exportToIcs(singleCall = null) {
  const calls = singleCall ? [singleCall] : getTargetCallsForExport();
  if (calls.length === 0) {
    showToast('No calls available to export!', 'error');
    return;
  }

  const titleTemplate = DOM.eventTitleInput.value.trim() || 'Call with {contact}';
  const nowStr = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';

  const icsLines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//OmniConverter//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    `X-WR-CALNAME:${titleTemplate}`
  ];

  calls.forEach(c => {
    if (!c.iso) return;
    try {
      const dtStart = new Date(c.iso);
      if (isNaN(dtStart.getTime())) return;

      const durCalendarSec = c.durationSec > 0 ? c.durationSec : 300;
      const dtEnd = new Date(dtStart.getTime() + (durCalendarSec * 1000));

      const dtstartStr = dtStart.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
      const dtendStr = dtEnd.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';

      let summary = titleTemplate;
      const contactName = (AppState.rawDataset && AppState.rawDataset.contact && AppState.rawDataset.contact.name) || 'Contact';
      summary = summary.replace('{contact}', contactName);
      summary = summary.replace('{type}', c.mediaType);
      summary = summary.replace('{duration}', c.durationFormatted);

      const description = `Type: ${c.mediaType} Call\\nDuration: ${c.durationFormatted}\\nStatus: ${c.status}\\nDirection: ${c.direction}`;

      icsLines.push("BEGIN:VEVENT");
      icsLines.push(`UID:${c.id}`);
      icsLines.push(`DTSTAMP:${nowStr}`);
      icsLines.push(`DTSTART:${dtstartStr}`);
      icsLines.push(`DTEND:${dtendStr}`);
      icsLines.push(`SUMMARY:${summary}`);
      icsLines.push(`DESCRIPTION:${description}`);
      icsLines.push("STATUS:CONFIRMED");
      icsLines.push("TRANSP:OPAQUE");
      icsLines.push("END:VEVENT");
    } catch (e) {}
  });

  icsLines.push("END:VCALENDAR");

  const icsContent = icsLines.join('\r\n');
  const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
  const fname = calls.length === 1 ? `call_${calls[0].date}_${calls[0].time.replace(/:/g, '')}.ics` : (calls.length === AppState.activeCalls.length ? 'call_logs.ics' : `call_logs_${calls.length}_selected.ics`);
  downloadBlob(blob, fname);
  showToast(`Calendar (.ics) generated successfully (${calls.length} event${calls.length > 1 ? 's' : ''})!`, 'success');
}

// 5. ZIP Bundle Exporter (All 4 formats)
async function exportToZip() {
  const calls = getTargetCallsForExport();
  if (calls.length === 0) {
    showToast('No calls available to export!', 'error');
    return;
  }

  showToast(`Preparing ZIP package with ${calls.length} calls in all 4 formats...`, 'info');

  const zip = new JSZip();

  // Excel
  const rows = calls.map((c, i) => ({
    "No.": i + 1,
    "Date": c.date,
    "Time": c.time,
    "Type": c.mediaType,
    "Direction": c.direction,
    "Status": c.status,
    "Duration (s)": c.durationSec,
    "Formatted Duration": c.durationFormatted,
    "UTC Start (ISO)": c.iso,
    "Record ID": c.id
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Calls");
  const xlsxBuf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  zip.file("call_logs.xlsx", xlsxBuf);

  // CSV
  const delimiter = ';';
  const csvHeaders = ["ID", "Date", "Time", "ISO_UTC", "Media_Type", "Direction", "Status", "Missed", "Duration_Seconds", "Duration_Formatted"];
  const csvLines = [csvHeaders.join(delimiter)];
  calls.forEach(c => {
    csvLines.push([`"${c.id}"`, `"${c.date}"`, `"${c.time}"`, `"${c.iso}"`, `"${c.mediaType}"`, `"${c.direction}"`, `"${c.status}"`, `"${c.isMissed ? 'Yes' : 'No'}"`, c.durationSec, `"${c.durationFormatted}"`].join(delimiter));
  });
  zip.file("call_logs.csv", '\uFEFF' + csvLines.join('\r\n'));

  // JSON
  zip.file("call_logs.json", JSON.stringify({ totalCalls: calls.length, calls: calls }, null, 2));

  // ICS
  const titleTemplate = DOM.eventTitleInput.value.trim() || 'Call with {contact}';
  const nowStr = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
  const icsLines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//OmniConverter//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"];
  calls.forEach(c => {
    if (!c.iso) return;
    try {
      const dtStart = new Date(c.iso);
      const dtEnd = new Date(dtStart.getTime() + ((c.durationSec || 300) * 1000));
      icsLines.push("BEGIN:VEVENT");
      icsLines.push(`UID:${c.id}`);
      icsLines.push(`DTSTAMP:${nowStr}`);
      icsLines.push(`DTSTART:${dtStart.toISOString().replace(/[-:]/g, '').split('.')[0]}Z`);
      icsLines.push(`DTEND:${dtEnd.toISOString().replace(/[-:]/g, '').split('.')[0]}Z`);
      icsLines.push(`SUMMARY:${titleTemplate}`);
      icsLines.push("STATUS:CONFIRMED");
      icsLines.push("END:VEVENT");
    } catch (e) {}
  });
  icsLines.push("END:VCALENDAR");
  zip.file("call_logs.ics", icsLines.join('\r\n'));

  const zipContent = await zip.generateAsync({ type: 'blob' });
  const fname = calls.length === AppState.activeCalls.length ? 'call_logs_bundle.zip' : `call_logs_${calls.length}_selected_bundle.zip`;
  downloadBlob(zipContent, fname);
  showToast('Complete ZIP bundle downloaded successfully!', 'success');
}

// Byte Formatter
function formatBytes(bytes) {
  if (!bytes || isNaN(bytes) || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

// Navigation Tab Switcher
function switchTab(tab) {
  if (tab === 'converter') {
    DOM.tabBtnConverter.classList.add('active');
    DOM.tabBtnSlicer.classList.remove('active');
    DOM.viewConverter.classList.remove('hidden');
    DOM.viewSlicer.classList.add('hidden');
  } else if (tab === 'slicer') {
    DOM.tabBtnSlicer.classList.add('active');
    DOM.tabBtnConverter.classList.remove('active');
    DOM.viewSlicer.classList.remove('hidden');
    DOM.viewConverter.classList.add('hidden');
  }
}

// ICS Parser Engine
function parseIcsCalendar(content) {
  const lines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const headers = [];
  const timezones = [];
  const events = [];

  let inVCalendar = false;
  let currentBlock = null;
  let blockLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    if (trimmed.toUpperCase() === 'BEGIN:VCALENDAR') {
      inVCalendar = true;
      headers.push(line);
      continue;
    }

    if (trimmed.toUpperCase() === 'END:VCALENDAR') {
      continue;
    }

    if (trimmed.toUpperCase().startsWith('BEGIN:')) {
      const compName = trimmed.substring(6).toUpperCase();
      if (compName === 'VEVENT' || compName === 'VTIMEZONE') {
        currentBlock = compName;
        blockLines = [line];
        continue;
      }
    }

    if (currentBlock) {
      blockLines.push(line);
      if (trimmed.toUpperCase() === `END:${currentBlock}`) {
        const blockText = blockLines.join('\r\n');
        if (currentBlock === 'VEVENT') {
          events.push(blockText);
        } else if (currentBlock === 'VTIMEZONE') {
          timezones.push(blockText);
        }
        currentBlock = null;
        blockLines = [];
      }
      continue;
    }

    if (inVCalendar && !currentBlock) {
      headers.push(line);
    }
  }

  if (headers.length === 0) {
    headers.push('BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//OmniConverter//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH');
  }

  return { headers, timezones, events };
}

// ICS Split Engine
function calculateIcsSplits() {
  if (SlicerState.events.length === 0) {
    SlicerState.parts = [];
    renderSlicerResults();
    return [];
  }

  const encoder = new TextEncoder();
  const headerBlock = SlicerState.headers.join('\r\n') + '\r\n';
  const tzBlock = SlicerState.timezones.length > 0 ? SlicerState.timezones.join('\r\n') + '\r\n' : '';
  const footerBlock = 'END:VCALENDAR\r\n';

  const baseHeaderFooter = headerBlock + tzBlock + footerBlock;
  const baseOverheadBytes = encoder.encode(baseHeaderFooter).length;

  const parts = [];
  const baseName = SlicerState.baseName.trim() || 'calendar_part';

  if (SlicerState.splitMode === 'size') {
    const maxBytes = Math.max(10240, SlicerState.targetSizeKB * 1024);
    let currentPartEvents = [];
    let currentBytes = baseOverheadBytes;

    for (let i = 0; i < SlicerState.events.length; i++) {
      const evt = SlicerState.events[i];
      const evtBytes = encoder.encode(evt + '\r\n').length;

      // Check if adding this event exceeds maxBytes
      if (currentPartEvents.length > 0 && (currentBytes + evtBytes) > maxBytes) {
        const partNum = parts.length + 1;
        const partContent = headerBlock + tzBlock + currentPartEvents.join('\r\n') + '\r\n' + footerBlock;
        const actualSizeBytes = encoder.encode(partContent).length;
        parts.push({
          partNum,
          filename: `${baseName}_${partNum}.ics`,
          eventsCount: currentPartEvents.length,
          sizeBytes: actualSizeBytes,
          content: partContent
        });

        currentPartEvents = [];
        currentBytes = baseOverheadBytes;
      }

      currentPartEvents.push(evt);
      currentBytes += evtBytes;
    }

    if (currentPartEvents.length > 0) {
      const partNum = parts.length + 1;
      const partContent = headerBlock + tzBlock + currentPartEvents.join('\r\n') + '\r\n' + footerBlock;
      const actualSizeBytes = encoder.encode(partContent).length;
      parts.push({
        partNum,
        filename: `${baseName}_${partNum}.ics`,
        eventsCount: currentPartEvents.length,
        sizeBytes: actualSizeBytes,
        content: partContent
      });
    }
  } else {
    // Mode: by event count
    const maxEvents = Math.max(1, SlicerState.targetEvents);
    let partNum = 1;
    for (let i = 0; i < SlicerState.events.length; i += maxEvents) {
      const chunk = SlicerState.events.slice(i, i + maxEvents);
      const partContent = headerBlock + tzBlock + chunk.join('\r\n') + '\r\n' + footerBlock;
      const actualSizeBytes = encoder.encode(partContent).length;
      parts.push({
        partNum: partNum++,
        filename: `${baseName}_${parts.length + 1}.ics`,
        eventsCount: chunk.length,
        sizeBytes: actualSizeBytes,
        content: partContent
      });
    }
  }

  SlicerState.parts = parts;
  renderSlicerResults();
  return parts;
}

// Render Slicer Output Cards & Statistics
function renderSlicerResults() {
  const parts = SlicerState.parts;
  const totalEvents = SlicerState.events.length;
  const origSize = SlicerState.sizeBytes;

  DOM.slicerStatEvents.textContent = totalEvents.toLocaleString();
  DOM.slicerStatSize.textContent = formatBytes(origSize);
  DOM.slicerStatParts.textContent = parts.length;
  DOM.slicerStatTz.textContent = SlicerState.timezones.length;
  DOM.slicerPartsCount.textContent = parts.length;

  if (parts.length === 0) {
    DOM.slicerPlanSummary.textContent = 'No events found in this calendar file.';
    DOM.slicerPartsList.innerHTML = '<div class="empty-state">No events available to split.</div>';
    DOM.btnDownloadAllZip.disabled = true;
    return;
  }

  DOM.btnDownloadAllZip.disabled = false;
  DOM.slicerPlanSummary.textContent = `Split into ${parts.length} file${parts.length > 1 ? 's' : ''} • All parts are safe for Google Calendar import!`;

  let eventOffset = 1;
  DOM.slicerPartsList.innerHTML = parts.map((part, idx) => {
    const startEvt = eventOffset;
    const endEvt = eventOffset + part.eventsCount - 1;
    eventOffset = endEvt + 1;

    const sizeFormatted = formatBytes(part.sizeBytes);
    const isGoogleSafe = part.sizeBytes <= 1048576; // <= 1MB
    const safePill = isGoogleSafe
      ? `<span class="part-size-pill">✓ ${sizeFormatted} (Safe)</span>`
      : `<span class="part-size-pill" style="background: rgba(239, 68, 68, 0.15); color: #f87171; border-color: rgba(239, 68, 68, 0.3);">⚠️ ${sizeFormatted}</span>`;

    return `
      <div class="part-card" data-part-index="${idx}">
        <div class="part-card-left">
          <div class="part-badge-num">${part.partNum}</div>
          <div>
            <div class="part-name">
              <span>${part.filename}</span>
              ${safePill}
            </div>
            <div class="part-submeta">
              Contains ${part.eventsCount.toLocaleString()} events (Events #${startEvt} – #${endEvt})
            </div>
          </div>
        </div>
        <div class="part-card-right">
          <button class="btn btn-sm btn-glass btn-download-part" data-part-index="${idx}">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            <span>Download .ICS</span>
          </button>
        </div>
      </div>
    `;
  }).join('');
}

// Download Individual Sliced Part
function downloadIcsPart(partIndex) {
  const part = SlicerState.parts[partIndex];
  if (!part) return;

  const blob = new Blob([part.content], { type: 'text/calendar;charset=utf-8' });
  downloadBlob(blob, part.filename);
  showToast(`Downloaded ${part.filename}`, 'success');
}

// Download All Parts as ZIP
async function downloadAllPartsZip() {
  if (!SlicerState.parts || SlicerState.parts.length === 0) {
    showToast('No sliced parts to download.', 'error');
    return;
  }

  if (typeof JSZip === 'undefined') {
    showToast('JSZip library not loaded. Downloading parts individually.', 'error');
    SlicerState.parts.forEach((_, idx) => downloadIcsPart(idx));
    return;
  }

  showToast('Preparing ZIP package of calendar slices...', 'info', 1800);
  const zip = new JSZip();
  const folder = zip.folder("calendar_slices");

  SlicerState.parts.forEach(part => {
    folder.file(part.filename, part.content);
  });

  const instructionsText = 
`GOOGLE CALENDAR IMPORT INSTRUCTIONS:
===================================
1. Unzip this package to your computer.
2. Open Google Calendar (https://calendar.google.com).
3. Click the Settings Gear Icon (top right) -> Settings.
4. In the left menu, select "Import & Export".
5. Upload each file (${SlicerState.parts[0]?.filename} etc.) one by one.
6. Choose the destination calendar and click "Import".

All files were created by OmniConverter to strictly comply with Google's 1MB upload limit.
Enjoy!`;

  folder.file("IMPORT_INSTRUCTIONS.txt", instructionsText);

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  const base = SlicerState.baseName || 'calendar';
  downloadBlob(zipBlob, `${base}_split_${SlicerState.parts.length}_parts.zip`);
  showToast(`Successfully downloaded ZIP with ${SlicerState.parts.length} calendar parts!`, 'success');
}

// Handle Slicer File Upload
function handleSlicerFileUpload(file) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      const parsed = parseIcsCalendar(text);

      if (parsed.events.length === 0) {
        showToast('No VEVENT components found in this .ics file.', 'error');
        return;
      }

      SlicerState.rawContent = text;
      SlicerState.filename = file.name;
      SlicerState.sizeBytes = file.size || new TextEncoder().encode(text).length;
      SlicerState.headers = parsed.headers;
      SlicerState.timezones = parsed.timezones;
      SlicerState.events = parsed.events;

      // Suggest base prefix from file name
      const cleanName = file.name.replace(/\.[^/.]+$/, '').replace(/[^a-zA-Z0-9_-]/g, '_');
      SlicerState.baseName = `${cleanName}_part`;
      DOM.slicerBaseName.value = SlicerState.baseName;

      DOM.slicerFilenameDisplay.textContent = file.name;
      DOM.slicerDetailsDisplay.textContent = `${parsed.events.length.toLocaleString()} events • ${formatBytes(SlicerState.sizeBytes)} • ${parsed.timezones.length} timezones`;
      DOM.slicerFileBanner.classList.remove('hidden');
      DOM.slicerWorkspace.classList.remove('hidden');

      calculateIcsSplits();
      showToast(`Loaded ${parsed.events.length} events from ${file.name}!`, 'success');
    } catch (err) {
      showToast(`Error parsing ICS file: ${err.message}`, 'error');
    }
  };
  reader.readAsText(file);
}

// File Upload / Parsing for Main View
function handleFileUpload(file) {
  if (!file) return;

  // Smart detect .ics uploaded in main tab -> redirect to Slicer!
  if (file.name.toLowerCase().endsWith('.ics')) {
    switchTab('slicer');
    handleSlicerFileUpload(file);
    showToast(`Detected .ics calendar file. Switched to ICS Slicer!`, 'info');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      if (file.name.endsWith('.json')) {
        const parsed = JSON.parse(text);
        loadDataset(normalizeDataset(parsed), file.name);
      } else {
        showToast('Please upload a compatible .json file or switch to ICS Slicer.', 'error');
      }
    } catch (err) {
      showToast(`Error reading file: ${err.message}`, 'error');
    }
  };
  reader.readAsText(file);
}

// Event Listeners Initialization
function initEventListeners() {
  const dropZone = DOM.dropZone;
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
    });
  });

  dropZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFileUpload(files[0]);
  });

  dropZone.addEventListener('click', (e) => {
    if (e.target.closest('#btn-reselect-file') || !e.target.closest('.loaded-banner')) {
      DOM.fileInput.click();
    }
  });

  DOM.btnReselectFile.addEventListener('click', (e) => {
    e.stopPropagation();
    DOM.fileInput.click();
  });

  DOM.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFileUpload(e.target.files[0]);
  });

  DOM.btnLoadSample.addEventListener('click', checkBackend);

  document.querySelectorAll('.tag-helper').forEach(btn => {
    btn.addEventListener('click', () => {
      const tag = btn.getAttribute('data-insert');
      DOM.eventTitleInput.value += ` ${tag}`;
      DOM.eventTitleInput.focus();
    });
  });

  DOM.btnExportXlsx.addEventListener('click', exportToExcel);
  DOM.btnExportCsv.addEventListener('click', exportToCsv);
  DOM.btnExportIcs.addEventListener('click', () => exportToIcs());
  DOM.btnExportJson.addEventListener('click', exportToJson);
  DOM.btnExportZip.addEventListener('click', exportToZip);

  DOM.searchInput.addEventListener('input', (e) => {
    AppState.filterQuery = e.target.value;
    DOM.btnClearSearch.classList.toggle('hidden', !AppState.filterQuery);
    applyFilters();
  });

  DOM.btnClearSearch.addEventListener('click', () => {
    DOM.searchInput.value = '';
    AppState.filterQuery = '';
    DOM.btnClearSearch.classList.add('hidden');
    applyFilters();
  });

  DOM.filterDateFrom.addEventListener('change', (e) => {
    AppState.filterDateFrom = e.target.value;
    applyFilters();
  });

  DOM.filterDateTo.addEventListener('change', (e) => {
    AppState.filterDateTo = e.target.value;
    applyFilters();
  });

  document.querySelectorAll('.filter-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      AppState.filterType = pill.getAttribute('data-filter-type');
      applyFilters();
    });
  });

  document.querySelectorAll('.btn-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.btn-toggle').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      AppState.timelineMode = btn.getAttribute('data-chart-mode');
      updateCharts();
    });
  });

  DOM.selectPageSize.addEventListener('change', (e) => {
    AppState.pageSize = parseInt(e.target.value, 10);
    AppState.currentPage = 1;
    renderTable();
  });

  DOM.btnSelectAll.addEventListener('click', () => {
    AppState.filteredCalls.forEach(c => AppState.selectedIds.add(c.id));
    renderTable();
    updateSelectionUI();
    showToast(`Selected ${AppState.selectedIds.size} calls for export`, 'info');
  });

  DOM.btnSelectNone.addEventListener('click', () => {
    AppState.selectedIds.clear();
    renderTable();
    updateSelectionUI();
    showToast('Selection cleared. Will export all matching calls.', 'info');
  });

  DOM.chkSelectPage.addEventListener('change', (e) => {
    const checked = e.target.checked;
    const pageSize = AppState.pageSize;
    const startIdx = (AppState.currentPage - 1) * pageSize;
    const pageCalls = AppState.filteredCalls.slice(startIdx, startIdx + pageSize);

    pageCalls.forEach(c => {
      if (checked) AppState.selectedIds.add(c.id);
      else AppState.selectedIds.delete(c.id);
    });
    renderTable();
    updateSelectionUI();
  });

  DOM.callsTableBody.addEventListener('change', (e) => {
    if (e.target.classList.contains('row-checkbox')) {
      const id = e.target.getAttribute('data-id');
      if (e.target.checked) AppState.selectedIds.add(id);
      else AppState.selectedIds.delete(id);

      const tr = e.target.closest('tr');
      if (tr) tr.classList.toggle('row-selected', e.target.checked);
      updateSelectionUI();

      if (AppState.selectedIds.size > 0) {
        showToast(`${AppState.selectedIds.size} call(s) selected for export`, 'info', 1800);
      }
    }
  });

  DOM.callsTableBody.addEventListener('click', (e) => {
    const inspectBtn = e.target.closest('.btn-inspect-row');
    if (inspectBtn) {
      const id = inspectBtn.getAttribute('data-id');
      openDetailsModal(id);
      return;
    }

    const singleIcsBtn = e.target.closest('.btn-export-single-ics');
    if (singleIcsBtn) {
      const id = singleIcsBtn.getAttribute('data-id');
      const call = AppState.activeCalls.find(c => c.id === id);
      if (call) exportToIcs(call);
    }
  });

  document.querySelectorAll('th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (AppState.currentSort.column === col) {
        AppState.currentSort.order = AppState.currentSort.order === 'asc' ? 'desc' : 'asc';
      } else {
        AppState.currentSort.column = col;
        AppState.currentSort.order = 'desc';
      }

      document.querySelectorAll('th.sortable').forEach(t => {
        t.classList.remove('sorted-asc', 'sorted-desc');
      });
      th.classList.add(AppState.currentSort.order === 'asc' ? 'sorted-asc' : 'sorted-desc');

      sortCalls();
      renderTable();
    });
  });

  DOM.btnPageFirst.addEventListener('click', () => {
    AppState.currentPage = 1;
    renderTable();
  });

  DOM.btnPagePrev.addEventListener('click', () => {
    if (AppState.currentPage > 1) {
      AppState.currentPage--;
      renderTable();
    }
  });

  DOM.btnPageNext.addEventListener('click', () => {
    const totalPages = Math.ceil(AppState.filteredCalls.length / AppState.pageSize);
    if (AppState.currentPage < totalPages) {
      AppState.currentPage++;
      renderTable();
    }
  });

  DOM.btnPageLast.addEventListener('click', () => {
    AppState.currentPage = Math.ceil(AppState.filteredCalls.length / AppState.pageSize);
    renderTable();
  });

  DOM.btnCloseModal.addEventListener('click', closeModal);
  DOM.btnModalDone.addEventListener('click', closeModal);
  DOM.modalDetails.addEventListener('click', (e) => {
    if (e.target === DOM.modalDetails) closeModal();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !DOM.modalDetails.classList.contains('hidden')) {
      closeModal();
    }
  });

  // Navigation Tab Switching
  if (DOM.tabBtnConverter && DOM.tabBtnSlicer) {
    DOM.tabBtnConverter.addEventListener('click', () => switchTab('converter'));
    DOM.tabBtnSlicer.addEventListener('click', () => switchTab('slicer'));
  }

  // Slicer Drag & Drop Listeners
  if (DOM.slicerDropZone) {
    ['dragenter', 'dragover'].forEach(eventName => {
      DOM.slicerDropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        DOM.slicerDropZone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      DOM.slicerDropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        DOM.slicerDropZone.classList.remove('drag-over');
      });
    });

    DOM.slicerDropZone.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files.length > 0) handleSlicerFileUpload(files[0]);
    });

    DOM.slicerDropZone.addEventListener('click', (e) => {
      if (e.target.closest('#btn-slicer-reselect') || !e.target.closest('.loaded-banner')) {
        DOM.slicerFileInput.click();
      }
    });

    DOM.btnSlicerReselect.addEventListener('click', (e) => {
      e.stopPropagation();
      DOM.slicerFileInput.click();
    });

    DOM.slicerFileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleSlicerFileUpload(e.target.files[0]);
    });
  }

  // Slicer Mode Switchers (Size vs Events)
  if (DOM.btnModeSize && DOM.btnModeEvents) {
    DOM.btnModeSize.addEventListener('click', () => {
      DOM.btnModeSize.classList.add('active');
      DOM.btnModeEvents.classList.remove('active');
      DOM.blockSizeOptions.classList.remove('hidden');
      DOM.blockEventsOptions.classList.add('hidden');
      SlicerState.splitMode = 'size';
      calculateIcsSplits();
    });

    DOM.btnModeEvents.addEventListener('click', () => {
      DOM.btnModeEvents.classList.add('active');
      DOM.btnModeSize.classList.remove('active');
      DOM.blockEventsOptions.classList.remove('hidden');
      DOM.blockSizeOptions.classList.add('hidden');
      SlicerState.splitMode = 'events';
      calculateIcsSplits();
    });
  }

  // Slicer Size Presets
  if (DOM.sizePresetContainer) {
    DOM.sizePresetContainer.querySelectorAll('.preset-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        DOM.sizePresetContainer.querySelectorAll('.preset-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        const sizeVal = parseInt(pill.getAttribute('data-size'), 10);
        const unit = pill.getAttribute('data-unit') || 'KB';
        DOM.customSizeVal.value = sizeVal;
        DOM.customSizeUnit.value = unit;
        SlicerState.targetSizeKB = unit === 'MB' ? sizeVal * 1024 : sizeVal;
        calculateIcsSplits();
      });
    });

    DOM.customSizeVal.addEventListener('input', (e) => {
      DOM.sizePresetContainer.querySelectorAll('.preset-pill').forEach(p => p.classList.remove('active'));
      const val = parseInt(e.target.value, 10) || 950;
      const unit = DOM.customSizeUnit.value;
      SlicerState.targetSizeKB = unit === 'MB' ? val * 1024 : val;
      calculateIcsSplits();
    });

    DOM.customSizeUnit.addEventListener('change', (e) => {
      const val = parseInt(DOM.customSizeVal.value, 10) || 950;
      const unit = e.target.value;
      SlicerState.targetSizeKB = unit === 'MB' ? val * 1024 : val;
      calculateIcsSplits();
    });
  }

  // Slicer Events Presets
  if (DOM.eventsPresetContainer) {
    DOM.eventsPresetContainer.querySelectorAll('.preset-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        DOM.eventsPresetContainer.querySelectorAll('.preset-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        const eventsVal = parseInt(pill.getAttribute('data-events'), 10);
        DOM.customEventsVal.value = eventsVal;
        SlicerState.targetEvents = eventsVal;
        calculateIcsSplits();
      });
    });

    DOM.customEventsVal.addEventListener('input', (e) => {
      DOM.eventsPresetContainer.querySelectorAll('.preset-pill').forEach(p => p.classList.remove('active'));
      const val = parseInt(e.target.value, 10) || 500;
      SlicerState.targetEvents = val;
      calculateIcsSplits();
    });
  }

  // Base output name
  if (DOM.slicerBaseName) {
    DOM.slicerBaseName.addEventListener('input', (e) => {
      SlicerState.baseName = e.target.value.trim() || 'calendar_part';
      calculateIcsSplits();
    });
  }

  // Recalculate Split Button
  if (DOM.btnRecalculateSplit) {
    DOM.btnRecalculateSplit.addEventListener('click', () => {
      calculateIcsSplits();
      showToast('Split recalculated successfully!', 'info');
    });
  }

  // Batch Download as ZIP
  if (DOM.btnDownloadAllZip) {
    DOM.btnDownloadAllZip.addEventListener('click', downloadAllPartsZip);
  }

  // Download Individual Part from list
  if (DOM.slicerPartsList) {
    DOM.slicerPartsList.addEventListener('click', (e) => {
      const btn = e.target.closest('.btn-download-part');
      if (btn) {
        const idx = parseInt(btn.getAttribute('data-part-index'), 10);
        downloadIcsPart(idx);
      }
    });
  }
}

// Bootstrap
document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  checkBackend();
});
