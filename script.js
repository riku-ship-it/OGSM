const GAS_URL = 'https://script.google.com/macros/s/AKfycbywtjAiFsqzEzKiQ4soH7LdRHWiViQTCrte3fL2ySS49nPb_w3Zk7ctX1dyb1A0zDmMXw/exec';

// ── Color Map ──
const COLOR_MAP = {
  teal:   '#0D9373',
  blue:   '#2B7FE0',
  amber:  '#C47A15',
  coral:  '#D44E28',
  purple: '#5B4FCF',
};
const AVATAR_COLORS = ['#2B7FE0','#0D9373','#C47A15','#D44E28','#5B4FCF','#0F766E','#7C3AED'];
const avatarCache = {};
function avatarColor(name) {
  if (!avatarCache[name]) {
    avatarCache[name] = AVATAR_COLORS[Object.keys(avatarCache).length % AVATAR_COLORS.length];
  }
  return avatarCache[name];
}
function initials(name) { return name ? name.slice(0,2) : '?'; }

// ── State ──
let state = { objectives: [], goals: [], actions: [], strategies: [] };
let selectedGoalId        = null;
let selectedStrategy      = null;
let editingActionId       = null;
let addingToGoalId        = null;
let editingTrafficGoalId  = null;
let editingGoalId         = null;
let editingStrategyGoalId = null;
let editingStrategyName   = null;
let pendingDeleteFn       = null;

// ── Staff State ──
const staffDataCache = {};
let currentStaff = localStorage.getItem('ogsm-current-staff') || 'Riku';
let staffList    = [];

// ── Tab State ──
let currentTab = 'ogsm';
let statsWeekOffset = 0;
let meetingWeekOffset = 0;
let statsEditingId = null;
let weekNoteCache = {};
let weekNoteTimers = {};

const TYPE_SCORES = {
  '(小型)舊流程/規則優化': 1,
  '(小型)小功能修改': 1,
  '(中型)新機制建立': 5,
  '(中型)系統功能新增': 5,
  '(中型)系統發布推廣': 5,
  '(大型)重大系統改版': 10,
  '(大型)重大功能導入': 10,
  '(超大型)全新平台導入': 20,
  // legacy keys for backward compat
  '大型・新機制': 10,
  '中型・新機制': 5,
  '小型・新機制': 1,
  '大型・功能修改': 10,
  '中型・功能修改': 5,
  '小型・功能修改': 1,
};
const TYPE_OPTIONS = [
  '(小型)舊流程/規則優化',
  '(小型)小功能修改',
  '(中型)新機制建立',
  '(中型)系統功能新增',
  '(中型)系統發布推廣',
  '(大型)重大系統改版',
  '(大型)重大功能導入',
  '(超大型)全新平台導入',
];
const TARGET_OPTIONS = ['全公司','特定部門','內部協作','外部企業'];
const STAFF_AVATAR_COLORS = ['#185FA5','#0F6E56','#C47A15','#8B3CC4','#D44E28','#2B7FE0'];

function switchTab(tab) {
  currentTab = tab;
  document.getElementById('tab-content-ogsm').style.display = tab === 'ogsm' ? '' : 'none';
  document.getElementById('tab-content-stats').style.display = tab === 'stats' ? '' : 'none';
  document.getElementById('tab-btn-ogsm').classList.toggle('active', tab === 'ogsm');
  document.getElementById('tab-btn-stats').classList.toggle('active', tab === 'stats');
  if (tab === 'stats') loadStats();
  else loadAndRender();
}

// ── Stats Data ──
function getStatsData() {
  try { return JSON.parse(localStorage.getItem('ogsm-stats') || '{}'); }
  catch(e) { return {}; }
}
function saveStatsData(data) { localStorage.setItem('ogsm-stats', JSON.stringify(data)); }
function getPersonStats(person) { return getStatsData()[person] || []; }

async function loadStats() {
  renderStats();
  try {
    const res = await fetch(GAS_URL + '?api=1&action=get_stats&staff=' + encodeURIComponent(currentStaff) + '&_t=' + Date.now(), { method: 'GET', cache: 'no-store' });
    const data = await res.json();
    if (Array.isArray(data.items)) {
      const allData = getStatsData();
      const localItems = allData[currentStaff] || [];
      const backendIds = new Set(data.items.map(function(i) { return i.id; }));
      const pendingItems = localItems.filter(function(i) { return !backendIds.has(i.id); });
      const backendMapped = data.items.map(function(item) {
        return { id: item.id, launchDate: item.launchDate, platform: item.platform, target: item.target, description: item.description, type: item.type, score: item.score, date: item.launchDate };
      });
      allData[currentStaff] = backendMapped.concat(pendingItems);
      saveStatsData(allData);
      renderStats();
    }
  } catch(e) { /* silently use localStorage */ }
}

async function postStatsToBackend(payload) {
  try {
    await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ ...payload, staff: currentStaff }) });
  } catch(e) { /* silently fail */ }
}

function getWeekStart(offsetWeeks) {
  const d = new Date();
  const day = d.getDay();
  const diff = (day - 4 + 7) % 7;
  const thu = new Date(d);
  thu.setDate(d.getDate() - diff + offsetWeeks * 7);
  thu.setHours(0, 0, 0, 0);
  return thu;
}
function getWeekEnd(weekStart) {
  const d = new Date(weekStart);
  d.setDate(d.getDate() + 6);
  return d;
}
function fmtMD(date) { return (date.getMonth() + 1) + '/' + date.getDate(); }
function isoDate(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
}

function renderStats() {
  const weekStart = getWeekStart(statsWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const weekStartStr = isoDate(weekStart);
  const weekEndStr = isoDate(weekEnd);
  const weekRangeStr = weekStartStr + '~' + weekEndStr;

  const allPersonItems = getPersonStats(currentStaff);
  const personItems = allPersonItems.filter(function(i) {
    const d = i.launchDate || i.date;
    return d >= weekStartStr && d <= weekEndStr;
  });
  const totalScore = personItems.reduce(function(s, i) { return s + (i.score || 0); }, 0);
  const smallCount = personItems.filter(function(i) { return i.type && (i.type.startsWith('(小型)') || i.type.includes('小型・')); }).length;
  const mediumCount = personItems.filter(function(i) { return i.type && (i.type.startsWith('(中型)') || i.type.includes('中型・')); }).length;
  const largeCount = personItems.filter(function(i) { return i.type && (i.type.startsWith('(大型)') || i.type.startsWith('(超大型)') || i.type.includes('大型・')); }).length;

  const wrap = document.getElementById('tab-content-stats');
  const showAddForm = wrap.dataset.showForm === '1';

  const noteHtml = '<div class="stats-note-editor-wrap">' +
    '<div class="stats-note-toolbar">' +
      '<button class="stats-note-toolbar-btn" onmousedown="event.preventDefault();weekNoteCmd(\'bold\')" title="粗體"><b>B</b></button>' +
      '<button class="stats-note-toolbar-btn" onmousedown="event.preventDefault();weekNoteCmd(\'italic\')" title="斜體"><i>I</i></button>' +
      '<button class="stats-note-toolbar-btn" onmousedown="event.preventDefault();weekNoteCmd(\'link\')" title="超連結"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
      '<button class="stats-note-toolbar-btn" onmousedown="event.preventDefault();weekNoteCmd(\'list\')" title="列點"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4" cy="6" r="1.5" fill="currentColor" stroke="none"/><circle cx="4" cy="12" r="1.5" fill="currentColor" stroke="none"/><circle cx="4" cy="18" r="1.5" fill="currentColor" stroke="none"/></svg></button>' +
    '</div>' +
    '<div id="stats-note-editor" class="stats-note-editor" contenteditable="true" data-placeholder="記錄本週成果或發現的問題..."></div>' +
  '</div>';

  const itemsHtml = personItems.map(function(item) {
    if (statsEditingId === item.id) {
      const typeOpts = TYPE_OPTIONS.map(function(t) {
        return '<option value="' + escHtml(t) + '"' + (t === item.type ? ' selected' : '') + '>' + escHtml(t) + '</option>';
      }).join('');
      const tgOpts = TARGET_OPTIONS.map(function(t) {
        return '<option value="' + escHtml(t) + '"' + (t === (item.target||'') ? ' selected' : '') + '>' + escHtml(t) + '</option>';
      }).join('');
      return '<div class="stats-item-row stats-item-edit-row">' +
        '<input type="date" class="stats-form-input stats-form-date" id="ei-date" value="' + escHtml(item.launchDate || item.date || '') + '" />' +
        '<input type="text" class="stats-form-input" id="ei-platform" value="' + escHtml(item.platform || '') + '" placeholder="系統平台" />' +
        '<select class="stats-form-select" id="ei-target">' + tgOpts + '</select>' +
        '<div class="stats-desc-wrap" style="flex:2">' +
          '<div class="stats-desc-toolbar">' +
            '<button class="stats-desc-toolbar-btn" onmousedown="event.preventDefault();descLinkCmd(\'ei-desc\')" title="插入連結"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
          '</div>' +
          '<div id="ei-desc" class="stats-desc-editor" contenteditable="true" data-placeholder="項目說明"></div>' +
        '</div>' +
        '<select class="stats-form-select" id="ei-type" onchange="statsEditTypeChange()">' + typeOpts + '</select>' +
        '<input type="number" class="stats-form-input stats-form-score" id="ei-score" value="' + escHtml(String(item.score || '')) + '" />' +
        '<button class="stats-form-confirm" onclick="saveStatsItemEdit(\'' + escHtml(item.id) + '\')">儲存</button>' +
        '<button class="stats-form-cancel" onclick="cancelEditStatsItem()">取消</button>' +
        '</div>';
    }
    return '<div class="stats-item-row">' +
      '<div class="stats-item-date">' + escHtml(item.launchDate ? fmtDate(item.launchDate) : (item.date ? fmtDate(item.date) : '')) + '</div>' +
      '<div class="stats-platform-badge">' + escHtml(item.platform || '') + '</div>' +
      (item.target ? '<div class="stats-item-target">' + escHtml(item.target) + '</div>' : '<div class="stats-item-target"></div>') +
      '<div class="stats-item-desc">' + renderDescHtml(item.description || '') + '</div>' +
      '<div class="stats-item-type' + (item.type && item.type.startsWith('(超大型)') ? ' stats-item-type-xlarge' : item.type && item.type.startsWith('(大型)') ? ' stats-item-type-large' : item.type && item.type.startsWith('(中型)') ? ' stats-item-type-medium' : item.type && item.type.startsWith('(小型)') ? ' stats-item-type-small' : '') + '">' + escHtml(item.type || '') + '</div>' +
      '<div class="stats-item-score">+' + (item.score || 0) + '分</div>' +
      '<div class="stats-item-actions">' +
        '<button class="stats-item-edit-btn" onclick="startEditStatsItem(\'' + escHtml(item.id) + '\')">編輯</button>' +
        '<button class="stats-item-del-btn" onclick="deleteStatsItem(\'' + escHtml(item.id) + '\')">刪除</button>' +
      '</div>' +
      '</div>';
  }).join('');

  const typeOptions = TYPE_OPTIONS.map(function(t) {
    return '<option value="' + escHtml(t) + '">' + escHtml(t) + '</option>';
  }).join('');

  const targetOptsHtml = TARGET_OPTIONS.map(function(t) {
    return '<option value="' + escHtml(t) + '">' + escHtml(t) + '</option>';
  }).join('');
  const addFormHtml = showAddForm
    ? '<div class="stats-add-form" id="stats-add-form">' +
        '<input type="date" class="stats-form-input stats-form-date" id="sf-date" />' +
        '<input type="text" class="stats-form-input" id="sf-platform" placeholder="系統平台（如 BBP）" />' +
        '<select class="stats-form-select" id="sf-target">' + targetOptsHtml + '</select>' +
        '<div class="stats-desc-wrap" style="flex:2">' +
          '<div class="stats-desc-toolbar">' +
            '<button class="stats-desc-toolbar-btn" onmousedown="event.preventDefault();descLinkCmd(\'sf-desc\')" title="插入連結"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
          '</div>' +
          '<div id="sf-desc" class="stats-desc-editor" contenteditable="true" data-placeholder="項目說明"></div>' +
        '</div>' +
        '<select class="stats-form-select" id="sf-type" onchange="statsTypeChange()">' + typeOptions + '</select>' +
        '<input type="number" class="stats-form-input stats-form-score" id="sf-score" placeholder="分數" />' +
        '<button class="stats-form-confirm" onclick="confirmAddStatsItem()">確認</button>' +
        '<button class="stats-form-cancel" onclick="cancelAddStatsItem()">取消</button>' +
        '</div>'
    : '<button class="stats-add-btn" onclick="openAddStatsForm()">+ 新增上線項目</button>';

  wrap.innerHTML =
    '<div class="stats-two-col">' +
      '<div class="stats-left-col">' +
        '<div class="stats-score-card">' +
          '<div class="stats-score-label">' + escHtml(currentStaff) + ' 本週得分</div>' +
          '<div class="stats-score-value">' + totalScore + ' <span>分</span></div>' +
          '<div class="stats-week-range">' +
            '<button class="stats-week-nav" onclick="statsNavWeek(-1)">‹</button>' +
            '<span>' + fmtMD(weekStart) + ' – ' + fmtMD(weekEnd) + '</span>' +
            '<button class="stats-week-nav" onclick="statsNavWeek(1)">›</button>' +
          '</div>' +
          '<div class="stats-size-breakdown">' +
            '<div class="stats-size-row">' +
              '<div class="stats-size-card stats-size-small">' +
                '<div class="stats-size-num">' + smallCount + '</div>' +
                '<div class="stats-size-lbl">小型 ×' + smallCount + '</div>' +
              '</div>' +
              '<div class="stats-size-card stats-size-medium">' +
                '<div class="stats-size-num">' + mediumCount + '</div>' +
                '<div class="stats-size-lbl">中型 ×' + mediumCount + '</div>' +
              '</div>' +
            '</div>' +
            '<div class="stats-size-card stats-size-large">' +
              '<div class="stats-size-lbl">大型 / 超大型 ×' + largeCount + '</div>' +
            '</div>' +
          '</div>' +
        '</div>' +
        '<div class="stats-legend-card">' +
          '<div class="stats-legend-title">計分標準</div>' +
          '<div class="stats-legend-list">' +
            '<div class="stats-legend-row stats-legend-small"><span class="stats-legend-name">小型</span><span class="stats-legend-pts">1 分</span></div>' +
            '<div class="stats-legend-row stats-legend-medium"><span class="stats-legend-name">中型</span><span class="stats-legend-pts">5 分</span></div>' +
            '<div class="stats-legend-row stats-legend-large"><span class="stats-legend-name">大型</span><span class="stats-legend-pts">10 分</span></div>' +
            '<div class="stats-legend-row stats-legend-xlarge"><span class="stats-legend-name">超大型</span><span class="stats-legend-pts">20 分</span></div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      '<div class="stats-right-col">' +
        '<div class="stats-note-area">' +
          '<div class="stats-note-header">' +
            '<span class="stats-note-header-title">本週成果 / 發現問題</span>' +
            '<span id="stats-note-save-status" class="stats-note-save-status"></span>' +
          '</div>' +
          noteHtml +
        '</div>' +
        '<div class="stats-items-section">' +
          '<div class="stats-items-header">' +
            '<span class="stats-items-label">本週上線項目</span>' +
            '<span class="stats-items-count">共 ' + personItems.length + ' 筆</span>' +
          '</div>' +
          '<div class="stats-items-list">' + itemsHtml + '</div>' +
          addFormHtml +
        '</div>' +
      '</div>' +
    '</div>';

  if (showAddForm) {
    statsTypeChange();
    const dateEl = document.getElementById('sf-date');
    if (dateEl) dateEl.value = isoDate(new Date());
  }
  if (statsEditingId) {
    const editingItem = (getStatsData()[currentStaff] || []).find(function(i) { return i.id === statsEditingId; });
    const eiDesc = document.getElementById('ei-desc');
    if (eiDesc && editingItem) eiDesc.innerHTML = editingItem.description || '';
  }
  const editor = document.getElementById('stats-note-editor');
  if (editor) {
    editor.oninput = scheduleWeekNoteSave;
    editor.addEventListener('click', function(e) {
      if (e.target.tagName === 'A') { e.preventDefault(); window.open(e.target.href, '_blank'); }
    });
    initWeekNoteEditor(currentStaff, weekRangeStr);
  }
}

function statsNavWeek(dir) {
  statsWeekOffset += dir;
  renderStats();
}

function openAddStatsForm() {
  const wrap = document.getElementById('tab-content-stats');
  wrap.dataset.showForm = '1';
  renderStats();
}

function statsTypeChange() {
  const typeEl = document.getElementById('sf-type');
  const scoreEl = document.getElementById('sf-score');
  if (typeEl && scoreEl && !scoreEl.dataset.manual) {
    scoreEl.value = TYPE_SCORES[typeEl.value] || '';
  }
}

function cancelAddStatsItem() {
  document.getElementById('tab-content-stats').dataset.showForm = '0';
  renderStats();
}

async function initWeekNoteEditor(person, weekStartStr) {
  const cacheKey = person + '-' + weekStartStr;
  const el = document.getElementById('stats-note-editor');
  if (!el) return;
  if (weekNoteCache[cacheKey] !== undefined) {
    el.innerHTML = weekNoteCache[cacheKey];
    return;
  }
  try {
    const res = await fetch(GAS_URL + '?api=1&action=get_week_note&staff=' + encodeURIComponent(person) + '&weekStart=' + weekStartStr + '&_t=' + Date.now(), { cache: 'no-store' });
    const data = await res.json();
    weekNoteCache[cacheKey] = data.content || '';
  } catch(e) {
    weekNoteCache[cacheKey] = '';
  }
  if (currentStaff !== person || isoDate(getWeekStart(statsWeekOffset)) !== weekStartStr) return;
  const editor = document.getElementById('stats-note-editor');
  if (editor) editor.innerHTML = weekNoteCache[cacheKey];
}

function scheduleWeekNoteSave() {
  const editor = document.getElementById('stats-note-editor');
  if (!editor) return;
  const ws = getWeekStart(statsWeekOffset);
  const weekRangeStr = isoDate(ws) + '~' + isoDate(getWeekEnd(ws));
  const cacheKey = currentStaff + '-' + weekRangeStr;
  weekNoteCache[cacheKey] = editor.innerHTML;
  const person = currentStaff;
  clearTimeout(weekNoteTimers[cacheKey]);
  weekNoteTimers[cacheKey] = setTimeout(async function() {
    try {
      await fetch(GAS_URL, {
        method: 'POST',
        body: JSON.stringify({ type: 'save_week_note', staff: person, weekStart: weekRangeStr, content: weekNoteCache[cacheKey] || '' })
      });
      delete weekNoteTimers[cacheKey];
      const s = document.getElementById('stats-note-save-status');
      if (s) { s.textContent = '已儲存'; setTimeout(function() { if (s) s.textContent = ''; }, 2000); }
    } catch(e) { delete weekNoteTimers[cacheKey]; }
  }, 1500);
}

function weekNoteCmd(cmd) {
  const editor = document.getElementById('stats-note-editor');
  if (!editor) return;
  editor.focus();
  if (cmd === 'bold') {
    document.execCommand('bold', false, null);
  } else if (cmd === 'italic') {
    document.execCommand('italic', false, null);
  } else if (cmd === 'link') {
    const url = prompt('請輸入連結網址：');
    if (url) {
      const sel = window.getSelection();
      if (sel && sel.toString()) {
        document.execCommand('createLink', false, url);
      } else {
        const text = prompt('請輸入顯示文字（留空則用網址）：') || url;
        document.execCommand('insertHTML', false, '<a href="' + url + '">' + text + '</a>');
      }
    }
  } else if (cmd === 'list') {
    document.execCommand('insertUnorderedList', false, null);
  }
  scheduleWeekNoteSave();
}
function descLinkCmd(editorId) {
  const editor = document.getElementById(editorId);
  if (!editor) return;
  editor.focus();
  const url = prompt('請輸入連結網址：');
  if (!url) return;
  const sel = window.getSelection();
  if (sel && sel.toString()) {
    document.execCommand('createLink', false, url);
    editor.querySelectorAll('a[href="' + url + '"]').forEach(function(a) {
      a.target = '_blank'; a.rel = 'noopener';
    });
  } else {
    const text = prompt('請輸入顯示文字（留空則用網址）：') || url;
    document.execCommand('insertHTML', false, '<a href="' + url + '" target="_blank" rel="noopener">' + text + '</a>');
  }
}
function renderDescHtml(html) {
  if (!html) return '';
  if (!/<[a-z]/i.test(html)) return escHtml(html);
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  tmp.querySelectorAll('a').forEach(function(a) {
    const href = a.getAttribute('href') || '';
    if (/^https?:\/\//i.test(href)) {
      a.setAttribute('target', '_blank');
      a.setAttribute('rel', 'noopener');
    } else {
      a.replaceWith(document.createTextNode(a.textContent));
    }
  });
  tmp.querySelectorAll('*:not(a)').forEach(function(el) {
    el.replaceWith(document.createTextNode(el.textContent));
  });
  return tmp.innerHTML;
}
function startEditStatsItem(id) {
  statsEditingId = id;
  renderStats();
}
function cancelEditStatsItem() {
  statsEditingId = null;
  renderStats();
}
function statsEditTypeChange() {
  const typeEl = document.getElementById('ei-type');
  const scoreEl = document.getElementById('ei-score');
  if (typeEl && scoreEl) scoreEl.value = TYPE_SCORES[typeEl.value] || '';
}
function saveStatsItemEdit(id) {
  const launchDate = (document.getElementById('ei-date').value || '').trim();
  const platform = (document.getElementById('ei-platform').value || '').trim();
  const target = document.getElementById('ei-target').value;
  const eiDescEl = document.getElementById('ei-desc');
  const desc = eiDescEl ? eiDescEl.innerHTML.trim() : '';
  const descText = eiDescEl ? (eiDescEl.innerText || eiDescEl.textContent || '').trim() : '';
  const type = document.getElementById('ei-type').value;
  const score = parseInt(document.getElementById('ei-score').value) || TYPE_SCORES[type] || 0;
  if (!platform || !descText) { showToast('❌ 請填寫平台與項目說明', true); return; }
  const allData = getStatsData();
  const items = allData[currentStaff] || [];
  const idx = items.findIndex(function(i) { return i.id === id; });
  if (idx >= 0) {
    items[idx] = Object.assign({}, items[idx], { launchDate: launchDate, platform: platform, target: target, description: desc, type: type, score: score });
    allData[currentStaff] = items;
    saveStatsData(allData);
    postStatsToBackend({ type: 'update_stats_item', id: id, launchDate: launchDate, platform: platform, target: target, description: desc, type_name: type, score: score });
  }
  statsEditingId = null;
  renderStats();
  showToast('✅ 已更新');
}
function deleteStatsItem(id) {
  openConfirmDelete('確定要刪除此上線項目？此操作無法復原。', function() {
    const allData = getStatsData();
    if (allData[currentStaff]) {
      allData[currentStaff] = allData[currentStaff].filter(function(i) { return i.id !== id; });
      saveStatsData(allData);
    }
    postStatsToBackend({ type: 'delete_stats_item', id: id });
    renderStats();
    showToast('✅ 已刪除');
  });
}

function confirmAddStatsItem() {
  const launchDate = (document.getElementById('sf-date').value || '').trim();
  const platform = (document.getElementById('sf-platform').value || '').trim();
  const target = document.getElementById('sf-target').value;
  const sfDescEl = document.getElementById('sf-desc');
  const desc = sfDescEl ? sfDescEl.innerHTML.trim() : '';
  const descText = sfDescEl ? (sfDescEl.innerText || sfDescEl.textContent || '').trim() : '';
  const type = document.getElementById('sf-type').value;
  const scoreRaw = document.getElementById('sf-score').value;
  const score = parseInt(scoreRaw) || TYPE_SCORES[type] || 0;

  if (!platform || !descText) { showToast('❌ 請填寫平台與項目說明', true); return; }

  const weekStart = getWeekStart(statsWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const today = new Date();
  const clampedDate = today < weekStart ? weekStart : today > weekEnd ? weekEnd : today;

  const newId = Date.now().toString();
  const newLaunchDate = launchDate || isoDate(clampedDate);
  const allData = getStatsData();
  if (!allData[currentStaff]) allData[currentStaff] = [];
  allData[currentStaff].push({ id: newId, launchDate: newLaunchDate, platform: platform, target: target, description: desc, type: type, score: score, date: isoDate(clampedDate) });
  saveStatsData(allData);

  postStatsToBackend({ type: 'add_stats_item', id: newId, launchDate: newLaunchDate, platform: platform, target: target, description: desc, type_name: type, score: score });

  document.getElementById('tab-content-stats').dataset.showForm = '0';
  renderStats();
  showToast('✅ 上線項目已新增');
}

// ── Fetch / Post ──
async function fetchData(staff) {
  const res = await fetch(GAS_URL + '?api=1&staff=' + encodeURIComponent(staff || currentStaff) + '&_t=' + Date.now(), { method: 'GET', cache: 'no-store' });
  return await res.json();
}
async function fetchStaffList() {
  const res = await fetch(GAS_URL + '?api=1&action=staff_list', { method: 'GET' });
  const data = await res.json();
  return data.staff || [];
}
async function postData(payload) {
  const body = { ...payload, staff: currentStaff };
  const res = await fetch(GAS_URL, { method: 'POST', body: JSON.stringify(body) });
  return await res.json();
}

// ── Render ──
function render() {
  renderObjective();
  renderColumns();
}

function renderObjective() {
  const { objectives } = state;
  const obj = objectives[0] || { title: '請設定目的', id: '' };

  const section = document.getElementById('obj-section');
  section.style.display = 'block';

  const nameEl = document.getElementById('obj-name');
  if (nameEl.textContent !== obj.title) nameEl.textContent = obj.title;

  nameEl.onblur = async function() {
    const newTitle = nameEl.textContent.trim();
    if (!newTitle || newTitle === obj.title) { nameEl.textContent = obj.title; return; }
    try {
      if (!obj.id) {
        const res = await postData({ type:'create_objective', new_title:newTitle });
        if (res.success) { obj.id = res.obj_id; obj.title = newTitle; state.objectives[0] = obj; showToast('✅ 目的已建立'); }
        else { showToast('❌ '+(res.message||'建立失敗'), true); nameEl.textContent = obj.title; }
      } else {
        const res = await postData({ type:'rename_objective', obj_id:obj.id, new_title:newTitle });
        if (res.success) { obj.title = newTitle; state.objectives[0] = obj; showToast('✅ 目的已更新'); }
        else { showToast('❌ '+(res.message||'更新失敗'), true); nameEl.textContent = obj.title; }
      }
    } catch(e) { showToast('❌ 網路錯誤', true); nameEl.textContent = obj.title; }
  };
  let _objComposing = false;
  nameEl.oncompositionstart = function() { _objComposing = true; };
  nameEl.oncompositionend   = function() { _objComposing = false; };
  nameEl.onkeydown = function(e) {
    if (e.key==='Enter' && !_objComposing) { e.preventDefault(); nameEl.blur(); }
    if (e.key==='Escape') { nameEl.textContent = obj.title; nameEl.blur(); }
  };

}

function renderColumns() {
  const wrap = document.getElementById('three-col-wrap');
  wrap.innerHTML = '';

  // -- G column --
  const gCol = makeColumn('G', '支線目標', 'col-tag-g', state.goals.length);
  const gBody = gCol.querySelector('.col-body');

  if (!state.goals.length) {
    gBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">🎯</div><span>尚無支線目標</span></div>';
  } else {
    state.goals.forEach((goal, idx) => {
      const color = COLOR_MAP[goal.color] || COLOR_MAP.blue;
      const item = document.createElement('div');
      item.className = 'goal-item' + (selectedGoalId === goal.id ? ' active' : '');
      item.style.setProperty('--goal-color', color);
      const tl = goal.traffic_light || 'green';
      const tdefs = getTrafficDefs(goal.id);
      const tlLabel = tdefs[tl] || (tl === 'red' ? '紅燈' : tl === 'yellow' ? '黃燈' : '綠燈');
      const deadline = getGoalDeadline(goal.id);
      item.dataset.dragId = goal.id;
      item.innerHTML = `
        <div class="goal-item-top-row">
          <div style="display:flex;align-items:center;gap:4px">
            <span class="drag-handle" title="拖移排序">⠿</span>
            <div class="goal-item-num">目標 ${idx+1}</div>
          </div>
          <div class="goal-traffic-badge-wrap">
            <span class="traffic-badge traffic-badge-${escHtml(tl)}">
              <span class="traffic-badge-dot traffic-badge-dot-${escHtml(tl)}"></span>
              <span class="traffic-badge-label">${escHtml(tlLabel)}</span>
            </span>
            <button class="traffic-def-btn" title="定義燈號意思">✎</button>
          </div>
        </div>
        <div class="goal-item-name" contenteditable="true" spellcheck="false">${escHtml(goal.name)}</div>
        <div class="goal-item-meta">
          <span></span>
          <div class="goal-meta-right">
            <span class="goal-deadline-btn ${deadline ? 'has-date' : ''}">${deadline ? fmtDate(deadline) : '+ 截止日'}</span>
            <span class="goal-item-arrow">→</span>
          </div>
        </div>
      `;
      // Click to select
      item.addEventListener('click', function(e) {
        if (e.target.classList.contains('goal-item-name')) return;
        if (e.target.closest('.goal-traffic-badge-wrap')) return;
        selectedGoalId = selectedGoalId === goal.id ? null : goal.id;
        selectedStrategy = null;
        renderColumns();
      });
      // Traffic badge click → popup
      item.querySelector('.traffic-badge').addEventListener('click', function(e) {
        e.stopPropagation();
        closeTrafficPopup();
        const rect = this.getBoundingClientRect();
        const popup = document.createElement('div');
        popup.id = 'traffic-popup-fl';
        popup.className = 'traffic-popup-fl';
        popup.style.top  = (rect.bottom + 4) + 'px';
        popup.style.right = (window.innerWidth - rect.right) + 'px';
        const currentDefs = getTrafficDefs(goal.id);
        const lights = [
          { key: 'green',  label: currentDefs.green  || '綠燈' },
          { key: 'yellow', label: currentDefs.yellow || '黃燈' },
          { key: 'red',    label: currentDefs.red    || '紅燈' },
        ];
        popup.innerHTML = lights.map(l => `
          <div class="traffic-popup-opt${tl === l.key ? ' current' : ''}" data-light="${escHtml(l.key)}">
            <span class="traffic-popup-dot traffic-popup-dot-${escHtml(l.key)}"></span>
            <span>${escHtml(l.label)}</span>
          </div>
        `).join('') + `
          <div class="traffic-popup-sep"></div>
          <div class="traffic-popup-edit"><span>✎ 編輯定義</span></div>
        `;
        popup.querySelectorAll('.traffic-popup-opt').forEach(opt => {
          opt.addEventListener('click', function(ev) {
            ev.stopPropagation();
            closeTrafficPopup();
            updateGoalTraffic(goal.id, opt.dataset.light);
          });
        });
        popup.querySelector('.traffic-popup-edit').addEventListener('click', function(ev) {
          ev.stopPropagation();
          closeTrafficPopup();
          openTrafficDefModal(goal.id);
        });
        document.body.appendChild(popup);
      });
      // Define button
      item.querySelector('.traffic-def-btn').addEventListener('click', function(e) {
        e.stopPropagation();
        openTrafficDefModal(goal.id);
      });
      // Deadline button
      item.querySelector('.goal-deadline-btn').addEventListener('click', function(e) {
        e.stopPropagation();
        closeDeadlinePopup();
        showGoalDeadlinePopup(e.currentTarget, goal.id);
      });
      // Inline edit goal name
      const nameEl = item.querySelector('.goal-item-name');
      nameEl.addEventListener('blur', async function() {
        const newName = nameEl.textContent.trim();
        if (!newName || newName === goal.name) { nameEl.textContent = goal.name; return; }
        try {
          const res = await postData({ type:'rename_goal', goal_id:goal.id, new_name:newName });
          if (res.success) { goal.name = newName; showToast('✅ 目標名稱已更新'); }
          else { showToast('❌ '+(res.message||'更新失敗'), true); nameEl.textContent = goal.name; }
        } catch(e) { showToast('❌ 網路錯誤', true); nameEl.textContent = goal.name; }
      });
      let _goalComposing = false;
      nameEl.addEventListener('compositionstart', function() { _goalComposing = true; });
      nameEl.addEventListener('compositionend',   function() { _goalComposing = false; });
      nameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !_goalComposing) { e.preventDefault(); nameEl.blur(); }
        if (e.key==='Escape') { nameEl.textContent = goal.name; nameEl.blur(); }
      });
      item.addEventListener('contextmenu', function(e) {
        e.preventDefault();
        e.stopPropagation();
        showContextMenu(e.clientX, e.clientY, [
          { label: '編輯目標', icon: '✏️', action: () => openEditGoalModal(goal.id) },
          { label: '刪除目標', icon: '🗑', danger: true, action: () => {
            openConfirmDelete(
              `確定要刪除目標「${goal.name}」？\n此操作將一併刪除所有策略與行動，無法復原。`,
              () => deleteGoal(goal.id)
            );
          }}
        ]);
      });
      gBody.appendChild(item);
    });
  }

  setupDragDrop(gBody, onGoalsDrop);
  // Add goal btn
  const gAddBtn = document.createElement('button');
  gAddBtn.className = 'btn-col-add';
  gAddBtn.textContent = '+ 新增目標';
  gAddBtn.onclick = openAddGoalModal;
  gCol.appendChild(gAddBtn);

  wrap.appendChild(gCol);

  // -- S column --
  const selectedGoal = state.goals.find(g => g.id === selectedGoalId);
  const goalActions = selectedGoalId ? state.actions.filter(a => a.goal_id === selectedGoalId) : [];
  const strategies = selectedGoalId
    ? [...new Set(goalActions.map(a => a.strategy_name || '（未分類）'))]
    : [];

  const sCount = selectedGoalId ? strategies.length : '';
  const sCol = makeColumn('S', '策略', 'col-tag-s', sCount);
  const sBody = sCol.querySelector('.col-body');

  if (!selectedGoalId) {
    sBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">←</div><span>請先選擇左側目標</span></div>';
  } else if (!strategies.length) {
    sBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">📋</div><span>此目標尚無策略</span></div>';
  } else {
    strategies.forEach((strat, idx) => {
      const acts = goalActions.filter(a => (a.strategy_name||'（未分類）') === strat && a.action_name);
      const completedActs = acts.filter(a => a.status === '完成').length;
      const stratPct = acts.length > 0 ? Math.round(completedActs / acts.length * 100) : 0;
      const stratColor = COLOR_MAP[selectedGoal?.color] || COLOR_MAP.blue;
      const successDef = getStrategySuccessDef(selectedGoalId, strat);
      const stratStatus = getStrategyStatus(selectedGoalId, strat);
      const item = document.createElement('div');
      item.className = 'strategy-item' + (selectedStrategy === strat ? ' active' : '');
      item.style.setProperty('--goal-color', stratColor);
      item.dataset.dragId = strat;
      item.innerHTML = `
        <div class="strategy-item-top-row">
          <div style="display:flex;align-items:center;gap:6px">
            <span class="drag-handle" title="拖移排序">⠿</span>
            <div class="strategy-item-num">S${idx+1}</div>
            <span class="strategy-status-badge strategy-status-badge-${escHtml(stratStatus)}">
              <span class="strategy-status-dot strategy-status-dot-${escHtml(stratStatus)}"></span>
              <span>${escHtml(stratStatus)}</span>
            </span>
          </div>
          <span class="strategy-pct-badge" style="color:${stratColor};border-color:color-mix(in srgb,${stratColor} 50%,transparent)">${stratPct}%</span>
        </div>
        <div class="strategy-item-name" contenteditable="true" spellcheck="false">${escHtml(strat)}</div>
        ${successDef ? `<div class="strategy-success-def">${escHtml(successDef)}</div>` : ''}
        <div style="display:flex;align-items:center;justify-content:space-between;gap:8px">
          <div class="strategy-progress-wrap">
            <div class="strategy-mini-bar"><div class="strategy-mini-fill" style="width:${stratPct}%;background:${stratColor}"></div></div>
            <span class="strategy-pct" style="color:${stratColor}">${completedActs}/${acts.length}</span>
          </div>
          <span class="strategy-item-arrow">→</span>
        </div>
      `;
      item.addEventListener('click', function(e) {
        if (e.target.classList.contains('strategy-item-name')) return;
        if (e.target.closest('.strategy-status-badge')) return;
        selectedStrategy = selectedStrategy === strat ? null : strat;
        renderColumns();
      });
      // Strategy status badge click
      item.querySelector('.strategy-status-badge').addEventListener('click', function(e) {
        e.stopPropagation();
        closeStrategyStatusPopup();
        const rect = this.getBoundingClientRect();
        const popup = document.createElement('div');
        popup.id = 'strategy-status-popup-fl';
        popup.className = 'strategy-status-popup-fl';
        const statuses = ['未開始','進行中','完成','卡關'];
        popup.innerHTML = statuses.map(s => `
          <div class="strategy-status-opt${stratStatus === s ? ' current' : ''}" data-status="${escHtml(s)}">${escHtml(s)}</div>
        `).join('');
        popup.querySelectorAll('.strategy-status-opt').forEach(opt => {
          opt.addEventListener('click', function(ev) {
            ev.stopPropagation();
            closeStrategyStatusPopup();
            updateStrategyStatus(selectedGoalId, strat, opt.dataset.status);
          });
        });
        document.body.appendChild(popup);
        const pr = popup.getBoundingClientRect();
        let top  = rect.bottom + 4;
        let left = rect.left;
        if (left + pr.width > window.innerWidth) left = window.innerWidth - pr.width - 8;
        if (top  + pr.height > window.innerHeight) top = rect.top - pr.height - 4;
        popup.style.top  = top  + 'px';
        popup.style.left = left + 'px';
      });
      // Inline edit strategy name
      const sNameEl = item.querySelector('.strategy-item-name');
      sNameEl.addEventListener('blur', async function() {
        const newName = sNameEl.textContent.trim();
        if (!newName || newName === strat) { sNameEl.textContent = strat; return; }
        try {
          const res = await postData({ type:'rename_strategy', goal_id:selectedGoalId, old_name:strat, new_name:newName });
          if (res.success) {
            state.actions.forEach(a => { if (a.goal_id === selectedGoalId && (a.strategy_name||'（未分類）') === strat) a.strategy_name = newName; });
            state.strategies.forEach(s => { if (s.goal_id === selectedGoalId && s.name === strat) s.name = newName; });
            if (selectedStrategy === strat) selectedStrategy = newName;
            const oldDef = getStrategySuccessDef(selectedGoalId, strat);
            if (oldDef) { saveStrategySuccessDef(selectedGoalId, newName, oldDef); saveStrategySuccessDef(selectedGoalId, strat, ''); }
            showToast('✅ 策略名稱已更新');
            renderColumns();
          } else { showToast('❌ '+(res.message||'更新失敗'), true); sNameEl.textContent = strat; }
        } catch(e) { showToast('❌ 網路錯誤', true); sNameEl.textContent = strat; }
      });
      let _stratComposing = false;
      sNameEl.addEventListener('compositionstart', function() { _stratComposing = true; });
      sNameEl.addEventListener('compositionend',   function() { _stratComposing = false; });
      sNameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !_stratComposing) { e.preventDefault(); sNameEl.blur(); }
        if (e.key==='Escape') { sNameEl.textContent = strat; sNameEl.blur(); }
      });
      item.addEventListener('contextmenu', function(e) {
        e.preventDefault();
        e.stopPropagation();
        showContextMenu(e.clientX, e.clientY, [
          { label: '編輯策略', icon: '✏️', action: () => openEditStrategyModal(selectedGoalId, strat) },
          { label: '刪除策略', icon: '🗑', danger: true, action: () => {
            openConfirmDelete(
              `確定要刪除策略「${strat}」？\n此操作將一併刪除其下所有行動，無法復原。`,
              () => deleteStrategy(selectedGoalId, strat)
            );
          }}
        ]);
      });
      sBody.appendChild(item);
    });
    setupDragDrop(sBody, onStrategiesDrop);
  }

  // Add action btn for S
  if (selectedGoalId) {
    const sAddBtn = document.createElement('button');
    sAddBtn.className = 'btn-col-add';
    sAddBtn.textContent = '+ 新增策略';
    sAddBtn.onclick = () => openAddStrategyModal(selectedGoalId, selectedGoal?.name || '');
    sCol.appendChild(sAddBtn);
  }

  wrap.appendChild(sCol);

  // -- M column --
  const mActions = (selectedStrategy
    ? goalActions.filter(a => (a.strategy_name||'（未分類）') === selectedStrategy)
    : selectedGoalId ? goalActions : []).filter(a => a.action_name);

  const mLabel = selectedStrategy ? selectedStrategy : (selectedGoalId ? '全部行動' : '');
  const mCount = selectedGoalId ? mActions.length : '';
  const mCol = makeColumn('M', '行動項目', 'col-tag-m', mCount);
  const mBody = mCol.querySelector('.col-body');

  if (!selectedGoalId) {
    mBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">←</div><span>請先選擇左側目標</span></div>';
  } else if (!mActions.length) {
    mBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">📝</div><span>尚無行動項目</span></div>';
  } else {
    mActions.forEach(a => {
      const item = document.createElement('div');
      item.className = 'action-item';
      item.dataset.dragId = a.id;
      item.innerHTML = `
        <div class="action-item-top">
          <span class="drag-handle" title="拖移排序" style="margin-right:4px">⠿</span>
          <span class="action-item-name" contenteditable="true" spellcheck="false">${escHtml(a.action_name)}</span>
          <span class="action-badge badge-${a.status}">${escHtml(a.status)}</span>
        </div>
        <div class="action-item-meta">
          ${a.assignee ? `<span class="action-meta-assignee">
            <span class="avatar" style="background:${avatarColor(a.assignee)}">${initials(a.assignee)}</span>
            ${escHtml(a.assignee)}
          </span>` : '<span></span>'}
          ${a.due_date ? `<span class="action-meta-date">${fmtDate(a.due_date)}</span>` : ''}
        </div>
      `;
      item.addEventListener('click', function(e) {
        if (e.target.classList.contains('action-item-name')) return;
        if (e.target.classList.contains('action-badge')) return;
        openEditModal(a.id);
      });
      // Inline status change via badge click
      const badge = item.querySelector('.action-badge');
      badge.addEventListener('click', function(e) {
        e.stopPropagation();
        closeStatusPopup();
        const rect = badge.getBoundingClientRect();
        const popup = document.createElement('div');
        popup.id = 'action-status-popup-fl';
        popup.className = 'action-status-popup-fl';
        popup.style.top  = (rect.bottom + 4) + 'px';
        popup.style.right = (window.innerWidth - rect.right) + 'px';
        popup.innerHTML = `
          <div class="action-status-opt" data-status="未開始">未開始</div>
          <div class="action-status-opt" data-status="進行中">進行中</div>
          <div class="action-status-opt" data-status="完成">完成</div>
          <div class="action-status-opt" data-status="卡關">卡關</div>
        `;
        popup.querySelectorAll('.action-status-opt').forEach(opt => {
          opt.addEventListener('click', function(ev) {
            ev.stopPropagation();
            closeStatusPopup();
            updateActionStatus(a.id, opt.dataset.status);
          });
        });
        document.body.appendChild(popup);
      });
      // Inline edit action name
      const aNameEl = item.querySelector('.action-item-name');
      aNameEl.addEventListener('blur', async function() {
        const newName = aNameEl.textContent.trim();
        if (!newName || newName === a.action_name) { aNameEl.textContent = a.action_name; return; }
        try {
          const res = await postData({ type:'rename_action', action_id:a.id, new_name:newName });
          if (res.success) { a.action_name = newName; showToast('✅ 行動名稱已更新'); }
          else { showToast('❌ '+(res.message||'更新失敗'), true); aNameEl.textContent = a.action_name; }
        } catch(e) { showToast('❌ 網路錯誤', true); aNameEl.textContent = a.action_name; }
      });
      let _actComposing = false;
      aNameEl.addEventListener('compositionstart', function() { _actComposing = true; });
      aNameEl.addEventListener('compositionend',   function() { _actComposing = false; });
      aNameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !_actComposing) { e.preventDefault(); aNameEl.blur(); }
        if (e.key==='Escape') { aNameEl.textContent = a.action_name; aNameEl.blur(); }
      });
      item.addEventListener('contextmenu', function(e) {
        e.preventDefault();
        e.stopPropagation();
        showContextMenu(e.clientX, e.clientY, [
          { label: '刪除行動', icon: '🗑', danger: true, action: () => {
            openConfirmDelete(
              `確定要刪除行動「${a.action_name}」？\n此操作無法復原。`,
              () => deleteAction(a.id)
            );
          }}
        ]);
      });
      mBody.appendChild(item);
    });
    const _mRef = mActions.slice();
    setupDragDrop(mBody, function(o) { onActionsDrop(o, _mRef); });
  }

  if (selectedGoalId) {
    const mAddBtn = document.createElement('button');
    mAddBtn.className = 'btn-col-add';
    mAddBtn.textContent = '+ 新增行動項目';
    mAddBtn.onclick = () => openAddActionModal(selectedGoalId, selectedGoal?.name || '', selectedStrategy || '');
    mCol.appendChild(mAddBtn);
  }

  wrap.appendChild(mCol);
}

const COLUMN_TOOLTIPS = {
  O: `<div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">目的</span>
        <p class="ogsm-tooltip-desc">成功時，你預期會長什麼樣子？</p>
      </div>`,
  G: `<div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">目標名稱</span>
        <p class="ogsm-tooltip-desc">達成目的的量化里程碑，需含數字與日期。</p>
      </div>
      <div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">燈號定義</span>
        <p class="ogsm-tooltip-desc">紅／黃／綠各代表這個目標的什麼狀態（自定義）</p>
      </div>`,
  S: `<div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">策略名稱</span>
        <p class="ogsm-tooltip-desc">選擇用什麼方法達成目標？</p>
      </div>
      <div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">成功定義</span>
        <p class="ogsm-tooltip-desc">自定義執行到什麼狀態，這個策略才算完成？</p>
      </div>`,
  M: `<div class="ogsm-tooltip-section">
        <span class="ogsm-tooltip-label">行動項目</span>
        <p class="ogsm-tooltip-desc">推進策略的具體行動，指定負責人與截止日。</p>
      </div>`,
};

function initOgsmTooltips() {
  const tip = document.createElement('div');
  tip.id = 'ogsm-global-tip';
  tip.className = 'ogsm-tooltip';
  document.body.appendChild(tip);

  document.addEventListener('mouseover', e => {
    const wrap = e.target.closest('.ogsm-tooltip-wrap');
    if (!wrap) return;
    const key = wrap.dataset.tooltipKey;
    if (!key || !COLUMN_TOOLTIPS[key]) return;
    tip.innerHTML = COLUMN_TOOLTIPS[key];
    const rect = wrap.getBoundingClientRect();
    tip.style.top = (rect.bottom + 10) + 'px';
    tip.style.left = rect.left + 'px';
    tip.classList.add('ogsm-tooltip-visible');
  });

  document.addEventListener('mouseout', e => {
    const wrap = e.target.closest('.ogsm-tooltip-wrap');
    if (!wrap) return;
    tip.classList.remove('ogsm-tooltip-visible');
  });
}

function makeColumn(tag, title, tagClass, count) {
  const col = document.createElement('div');
  col.className = 'col-panel';
  col.innerHTML = `
    <div class="col-header">
      <div class="ogsm-tooltip-wrap" data-tooltip-key="${tag}">
        <span class="col-tag ${tagClass}">${tag}</span>
      </div>
      <span class="col-title">${title}</span>
      <span class="col-count">${count !== '' ? count + ' 項' : ''}</span>
    </div>
    <div class="col-body"></div>
  `;
  return col;
}

// ── Edit Modal ──
function openEditModal(actionId) {
  const a = state.actions.find(x => x.id === actionId);
  if (!a) return;
  editingActionId = actionId;
  document.getElementById('edit-modal-title').textContent = a.action_name;
  document.getElementById('edit-modal-sub').textContent   = a.strategy_name;
  document.getElementById('edit-progress').value = a.progress;
  document.getElementById('edit-progress-display').textContent = a.progress + '%';
  document.getElementById('edit-status').value    = a.status;
  document.getElementById('edit-assignee').value  = a.assignee || '';
  document.getElementById('edit-due-date').value  = a.due_date || '';
  document.getElementById('edit-action-name').value = a.action_name || '';
  const goalStratList = state.strategies.filter(s => s.goal_id === a.goal_id).map(s => s.name);
  const stratSel = document.getElementById('edit-action-strategy');
  stratSel.innerHTML = goalStratList.map(s => `<option value="${escHtml(s)}"${s === a.strategy_name ? ' selected' : ''}>${escHtml(s)}</option>`).join('');
  openOverlay('modal-edit');
}
function closeEditModal() { closeOverlay('modal-edit'); editingActionId = null; }
async function saveEditModal() {
  if (!editingActionId) return;
  const btn = document.getElementById('edit-save-btn');
  btn.disabled = true; btn.textContent = '儲存中';
  const payload = {
    type: 'update_action', id: editingActionId,
    progress: Number(document.getElementById('edit-progress').value),
    status:   document.getElementById('edit-status').value,
    assignee: document.getElementById('edit-assignee').value.trim(),
    due_date: document.getElementById('edit-due-date').value,
    action_name: document.getElementById('edit-action-name').value.trim(),
    strategy_name: document.getElementById('edit-action-strategy').value,
  };
  try {
    const res = await postData(payload);
    if (res.success) {
      const a = state.actions.find(x => x.id === editingActionId);
      if (a) {
        a.strategy_name = payload.strategy_name;
        a.action_name   = payload.action_name;
        a.progress      = payload.progress;
        a.status        = payload.status;
        a.assignee      = payload.assignee;
        a.due_date      = payload.due_date;
      }
      showToast('✅ 更新成功'); closeEditModal(); renderColumns();
    } else showToast('❌ '+(res.message||'更新失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '儲存更新'; }
}

// ── Add Goal Modal ──
function openAddGoalModal() {
  document.getElementById('new-goal-due').value = '';
  openOverlay('modal-add-goal');
}
function closeAddGoalModal() { closeOverlay('modal-add-goal'); }
async function saveNewGoal() {
  const name = document.getElementById('new-goal-name').value.trim();
  const color = document.getElementById('new-goal-color').value;
  const progress = Number(document.getElementById('new-goal-progress').value);
  const due = document.getElementById('new-goal-due').value;
  if (!name) { showToast('❌ 請填寫支線名稱', true); return; }
  const btn = document.getElementById('add-goal-save-btn');
  btn.disabled = true; btn.textContent = '新增中';
  const obj = state.objectives[0] || { id:'1' };
  const newGoalId = 'G'+Date.now();
  const payload = {
    type:'add_goal', obj_id:obj.id, obj_title:obj.title,
    goal_id:newGoalId, goal_name:name, goal_progress:progress, goal_color:color,
    action_id:'A'+Date.now(),
  };
  try {
    const res = await postData(payload);
    if (res.success) {
      if (due) saveGoalDeadline(newGoalId, due);
      showToast('✅ 新增成功'); closeAddGoalModal();
      document.getElementById('new-goal-name').value = '';
      document.getElementById('new-goal-progress').value = 0;
      document.getElementById('new-goal-progress-display').textContent = '0%';
      document.getElementById('new-goal-due').value = '';
      await loadAndRender();
    } else showToast('❌ '+(res.message||'新增失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '新增目標'; }
}

// ── Add Action Modal ──
function openAddActionModal(goalId, goalName, strategyName='') {
  addingToGoalId = goalId;
  document.getElementById('add-action-goal-name').textContent = goalName;
  document.getElementById('new-action-strategy').value = strategyName;
  document.getElementById('new-action-name').value = '';
  document.getElementById('new-action-assignee').value = '';
  document.getElementById('new-action-due').value = '';
  document.getElementById('new-action-progress').value = 0;
  document.getElementById('new-action-progress-display').textContent = '0%';
  document.getElementById('new-action-status').value = '未開始';
  openOverlay('modal-add-action');
}
function closeAddActionModal() { closeOverlay('modal-add-action'); addingToGoalId = null; }
async function saveNewAction() {
  const strategy = document.getElementById('new-action-strategy').value.trim();
  const name = document.getElementById('new-action-name').value.trim();
  if (!strategy || !name) { showToast('❌ 請填寫策略名稱與行動項目', true); return; }
  const btn = document.getElementById('add-action-save-btn');
  btn.disabled = true; btn.textContent = '新增中';
  const goal = state.goals.find(g => g.id === addingToGoalId);
  const obj = state.objectives[0] || { id:'1', title:'' };
  const payload = {
    type:'add_action', obj_id:obj.id, obj_title:obj.title,
    goal_id:addingToGoalId, goal_name:goal?.name||'', goal_progress:goal?.progress||0, goal_color:goal?.color||'blue',
    action_id:'A'+Date.now(), strategy_name:strategy, action_name:name,
    assignee:document.getElementById('new-action-assignee').value.trim(),
    due_date:document.getElementById('new-action-due').value,
    progress:Number(document.getElementById('new-action-progress').value),
    status:document.getElementById('new-action-status').value,
  };
  try {
    const res = await postData(payload);
    if (res.success) { showToast('✅ 新增行動成功'); closeAddActionModal(); await loadAndRender(); }
    else showToast('❌ '+(res.message||'新增失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '新增行動'; }
}

// ── Strategy Data helpers ──
function getStrategyData(goalId, stratName) {
  return state.strategies.find(s => s.goal_id === goalId && s.name === stratName) || { status: '', success_def: '' };
}
function getStrategySuccessDef(goalId, strategyName) {
  const sd = getStrategyData(goalId, strategyName);
  if (sd.success_def) return sd.success_def;
  // fallback to localStorage
  try { return JSON.parse(localStorage.getItem('ogsm-sdefs-' + goalId) || '{}')[strategyName] || ''; }
  catch(e) { return ''; }
}
function getStrategyStatus(goalId, stratName) {
  return getStrategyData(goalId, stratName).status || '未開始';
}
function saveStrategySuccessDef(goalId, strategyName, def) {
  // update in-memory state
  const sd = state.strategies.find(s => s.goal_id === goalId && s.name === strategyName);
  if (sd) sd.success_def = def;
  else state.strategies.push({ goal_id: goalId, name: strategyName, status: '', success_def: def });
  // keep localStorage for backward compat
  try {
    const defs = JSON.parse(localStorage.getItem('ogsm-sdefs-' + goalId) || '{}');
    if (def) defs[strategyName] = def; else delete defs[strategyName];
    localStorage.setItem('ogsm-sdefs-' + goalId, JSON.stringify(defs));
  } catch(e) {}
}
async function updateStrategyStatus(goalId, stratName, status) {
  const sd = state.strategies.find(s => s.goal_id === goalId && s.name === stratName);
  if (sd) sd.status = status;
  else state.strategies.push({ goal_id: goalId, name: stratName, status: status, success_def: '' });
  renderColumns();
  try {
    await postData({ type: 'update_strategy_status', goal_id: goalId, strategy_name: stratName, status: status });
  } catch(e) { showToast('❌ 網路錯誤', true); }
}

// ── Add Strategy Modal ──
function openAddStrategyModal(goalId, goalName) {
  addingToGoalId = goalId;
  document.getElementById('add-strategy-goal-name').textContent = goalName;
  document.getElementById('new-strategy-name').value = '';
  document.getElementById('new-strategy-success-def').value = '';
  openOverlay('modal-add-strategy');
}
function closeAddStrategyModal() { closeOverlay('modal-add-strategy'); addingToGoalId = null; }
async function saveNewStrategy() {
  const strategyName = document.getElementById('new-strategy-name').value.trim();
  if (!strategyName) { showToast('❌ 請填寫策略名稱', true); return; }
  const successDef = document.getElementById('new-strategy-success-def').value.trim();
  const btn = document.getElementById('add-strategy-save-btn');
  btn.disabled = true; btn.textContent = '新增中';
  const goal = state.goals.find(g => g.id === addingToGoalId);
  const obj = state.objectives[0] || { id:'1', title:'' };
  const actionId = 'A'+Date.now();
  const payload = {
    type:'add_action', obj_id:obj.id, obj_title:obj.title,
    goal_id:addingToGoalId, goal_name:goal?.name||'', goal_progress:goal?.progress||0, goal_color:goal?.color||'blue',
    action_id:actionId, strategy_name:strategyName, action_name:'',
    assignee:'', due_date:'', progress:0, status:'未開始',
  };
  try {
    const res = await postData(payload);
    if (res.success) {
      if (successDef) {
        saveStrategySuccessDef(addingToGoalId, strategyName, successDef);
        await postData({ type: 'update_action', id: actionId, success_def: successDef });
      }
      showToast('✅ 新增策略成功');
      closeAddStrategyModal();
      await loadAndRender();
    } else showToast('❌ '+(res.message||'新增失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '新增策略'; }
}

// ── Overlay helpers ──
function openOverlay(id)  { document.getElementById(id).classList.add('open'); }
function closeOverlay(id) { document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('.modal-overlay').forEach(el => {
  el.addEventListener('click', e => { if (e.target===el) el.classList.remove('open'); });
});
document.addEventListener('click', function() { closeStatusPopup(); closeTrafficPopup(); closeContextMenu(); closeDeadlinePopup(); closeStrategyStatusPopup(); });
document.addEventListener('contextmenu', function() { closeContextMenu(); });
document.addEventListener('keydown', e => {
  if (e.key==='Escape') {
    closeContextMenu();
    document.querySelectorAll('.modal-overlay.open').forEach(el => el.classList.remove('open'));
  }
});

// ── Toast ──
let toastTimer = null;
function showToast(msg, isError=false) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast'+(isError?' error':'');
  requestAnimationFrame(() => t.classList.add('show'));
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}

// ── Utils ──
function escHtml(str) {
  return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtDate(s) {
  if (!s) return '';
  const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return parseInt(m[2])+'/'+parseInt(m[3]);
  const d = new Date(s);
  return isNaN(d) ? '' : (d.getMonth()+1)+'/'+d.getDate();
}

// ── Goal Deadline helpers ──
function getGoalDeadline(goalId) {
  const goal = state.goals.find(g => g.id === goalId);
  return goal ? (goal.deadline || '') : '';
}
async function saveGoalDeadline(goalId, date) {
  const goal = state.goals.find(g => g.id === goalId);
  if (goal) goal.deadline = date;
  try {
    const res = await postData({ type: 'update_goal_deadline', goal_id: goalId, deadline: date });
    if (!res.success) showToast('❌ 日期儲存失敗', true);
  } catch(e) {
    showToast('❌ 網路錯誤', true);
  }
}
function closeDeadlinePopup() {
  const p = document.getElementById('goal-deadline-popup-fl');
  if (p) p.remove();
}
function showGoalDeadlinePopup(anchor, goalId) {
  const rect = anchor.getBoundingClientRect();
  const popup = document.createElement('div');
  popup.id = 'goal-deadline-popup-fl';
  popup.className = 'goal-deadline-popup-fl';
  const current = getGoalDeadline(goalId);
  popup.innerHTML = `
    <input type="date" class="deadline-popup-input" value="${escHtml(current)}" />
    <div class="deadline-popup-actions">
      <button class="deadline-popup-clear">清除</button>
      <button class="deadline-popup-save">確認</button>
    </div>
  `;
  // Position: prefer below the anchor, align left
  document.body.appendChild(popup);
  const popRect = popup.getBoundingClientRect();
  let top = rect.bottom + 4;
  let left = rect.left;
  if (top + popRect.height > window.innerHeight) top = rect.top - popRect.height - 4;
  if (left + popRect.width > window.innerWidth) left = window.innerWidth - popRect.width - 8;
  popup.style.top  = top + 'px';
  popup.style.left = left + 'px';
  popup.querySelector('.deadline-popup-save').addEventListener('click', function(e) {
    e.stopPropagation();
    const date = popup.querySelector('.deadline-popup-input').value;
    saveGoalDeadline(goalId, date);
    closeDeadlinePopup();
    renderColumns();
  });
  popup.querySelector('.deadline-popup-clear').addEventListener('click', function(e) {
    e.stopPropagation();
    saveGoalDeadline(goalId, '');
    closeDeadlinePopup();
    renderColumns();
  });
  popup.addEventListener('click', function(e) { e.stopPropagation(); });
  popup.querySelector('.deadline-popup-input').focus();
}

// ── Traffic Light helpers ──
function getTrafficDefs(goalId) {
  try { return JSON.parse(localStorage.getItem('ogsm-tdefs-' + goalId) || '{}'); }
  catch(e) { return {}; }
}
function saveTrafficDefs(goalId, defs) {
  localStorage.setItem('ogsm-tdefs-' + goalId, JSON.stringify(defs));
}
async function updateGoalTraffic(goalId, light) {
  const goal = state.goals.find(g => g.id === goalId);
  if (!goal) return;
  const prev = goal.traffic_light;
  goal.traffic_light = light;
  renderColumns();
  try {
    await postData({ type: 'update_goal_traffic', goal_id: goalId, traffic_light: light });
  } catch(e) {
    goal.traffic_light = prev;
    renderColumns();
    showToast('❌ 網路錯誤', true);
  }
}
function openTrafficDefModal(goalId) {
  editingTrafficGoalId = goalId;
  const goal = state.goals.find(g => g.id === goalId);
  document.getElementById('traffic-def-goal-name').textContent = goal ? goal.name : '';
  const defs = getTrafficDefs(goalId);
  document.getElementById('tdef-red').value    = defs.red    || '';
  document.getElementById('tdef-yellow').value = defs.yellow || '';
  document.getElementById('tdef-green').value  = defs.green  || '';
  openOverlay('modal-traffic-def');
}
function saveTrafficDefsUI() {
  if (!editingTrafficGoalId) return;
  saveTrafficDefs(editingTrafficGoalId, {
    red:    document.getElementById('tdef-red').value.trim(),
    yellow: document.getElementById('tdef-yellow').value.trim(),
    green:  document.getElementById('tdef-green').value.trim(),
  });
  closeOverlay('modal-traffic-def');
  renderColumns();
  showToast('✅ 燈號定義已儲存');
}

// ── Context Menu ──
function showContextMenu(x, y, items) {
  closeContextMenu();
  const menu = document.createElement('div');
  menu.id = 'context-menu-fl';
  menu.className = 'context-menu-fl';
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  items.forEach(item => {
    const div = document.createElement('div');
    div.className = 'context-menu-item' + (item.danger ? ' danger' : '');
    div.innerHTML = `<span class="context-menu-icon">${item.icon || ''}</span>${escHtml(item.label)}`;
    div.addEventListener('click', function(e) {
      e.stopPropagation();
      closeContextMenu();
      item.action();
    });
    menu.appendChild(div);
  });
  document.body.appendChild(menu);
  // Adjust if overflowing the viewport
  const rect = menu.getBoundingClientRect();
  if (rect.right  > window.innerWidth)  menu.style.left = (x - rect.width)  + 'px';
  if (rect.bottom > window.innerHeight) menu.style.top  = (y - rect.height) + 'px';
}
function closeContextMenu() {
  const m = document.getElementById('context-menu-fl');
  if (m) m.remove();
}
function openConfirmDelete(desc, onConfirm) {
  document.getElementById('confirm-delete-desc').textContent = desc;
  pendingDeleteFn = onConfirm;
  openOverlay('modal-confirm-delete');
}
function executeDelete() {
  closeOverlay('modal-confirm-delete');
  if (pendingDeleteFn) { pendingDeleteFn(); pendingDeleteFn = null; }
}

// ── Delete operations ──
async function deleteGoal(goalId) {
  try {
    const res = await postData({ type: 'delete_goal', goal_id: goalId });
    if (res.success) {
      showToast('✅ 目標已刪除');
      if (selectedGoalId === goalId) { selectedGoalId = null; selectedStrategy = null; }
      await loadAndRender();
    } else { showToast('❌ ' + (res.message || '刪除失敗'), true); }
  } catch(e) { showToast('❌ 網路錯誤', true); }
}
async function deleteStrategy(goalId, strategyName) {
  try {
    const res = await postData({ type: 'delete_strategy', goal_id: goalId, strategy_name: strategyName });
    if (res.success) {
      showToast('✅ 策略已刪除');
      if (selectedStrategy === strategyName) selectedStrategy = null;
      await loadAndRender();
    } else { showToast('❌ ' + (res.message || '刪除失敗'), true); }
  } catch(e) { showToast('❌ 網路錯誤', true); }
}
async function deleteAction(actionId) {
  try {
    const res = await postData({ type: 'delete_action', action_id: actionId });
    if (res.success) {
      showToast('✅ 行動已刪除');
      await loadAndRender();
    } else { showToast('❌ ' + (res.message || '刪除失敗'), true); }
  } catch(e) { showToast('❌ 網路錯誤', true); }
}

// ── Traffic popup ──
function closeTrafficPopup() {
  const p = document.getElementById('traffic-popup-fl');
  if (p) p.remove();
}

// ── Strategy Status popup ──
function closeStrategyStatusPopup() {
  const p = document.getElementById('strategy-status-popup-fl');
  if (p) p.remove();
}

// ── Edit Goal Modal ──
function openEditGoalModal(goalId) {
  editingGoalId = goalId;
  const goal = state.goals.find(g => g.id === goalId);
  document.getElementById('edit-goal-modal-sub').textContent = goal ? goal.name : '';
  document.getElementById('edit-goal-color').value = goal ? (goal.color || 'blue') : 'blue';
  openOverlay('modal-edit-goal');
}
async function saveEditGoal() {
  if (!editingGoalId) return;
  const color = document.getElementById('edit-goal-color').value;
  const btn = document.getElementById('edit-goal-save-btn');
  btn.disabled = true; btn.textContent = '儲存中';
  try {
    const res = await postData({ type: 'update_goal_color', goal_id: editingGoalId, color: color });
    if (res.success) {
      const goal = state.goals.find(g => g.id === editingGoalId);
      if (goal) goal.color = color;
      showToast('✅ 目標顏色已更新');
      closeOverlay('modal-edit-goal');
      renderColumns();
    } else showToast('❌ '+(res.message||'更新失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '儲存'; }
}

// ── Edit Strategy Modal ──
function openEditStrategyModal(goalId, stratName) {
  editingStrategyGoalId = goalId;
  editingStrategyName   = stratName;
  document.getElementById('edit-strategy-modal-sub').textContent = stratName;
  document.getElementById('edit-strategy-success-def').value = getStrategySuccessDef(goalId, stratName);
  openOverlay('modal-edit-strategy');
}
async function saveEditStrategy() {
  if (!editingStrategyGoalId || !editingStrategyName) return;
  const successDef = document.getElementById('edit-strategy-success-def').value.trim();
  const btn = document.getElementById('edit-strategy-save-btn');
  btn.disabled = true; btn.textContent = '儲存中';
  try {
    const placeholder = state.actions.find(a => a.goal_id === editingStrategyGoalId && a.strategy_name === editingStrategyName && !a.action_name);
    const res = placeholder?.id
      ? await postData({ type: 'update_action', id: placeholder.id, success_def: successDef })
      : await postData({ type: 'update_strategy_success_def', goal_id: editingStrategyGoalId, strategy_name: editingStrategyName, success_def: successDef });
    if (res.success) {
      saveStrategySuccessDef(editingStrategyGoalId, editingStrategyName, successDef);
      showToast('✅ 成功定義已更新');
      closeOverlay('modal-edit-strategy');
      renderColumns();
    } else showToast('❌ '+(res.message||'更新失敗'), true);
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '儲存'; }
}

// ── Action status inline update ──
function closeStatusPopup() {
  const p = document.getElementById('action-status-popup-fl');
  if (p) p.remove();
}
async function updateActionStatus(actionId, status) {
  const a = state.actions.find(x => x.id === actionId);
  if (!a) return;
  const prev = a.status;
  a.status = status;
  renderColumns();
  try {
    await postData({ type: 'update_action', id: actionId, status });
  } catch(e) {
    a.status = prev;
    renderColumns();
    showToast('❌ 網路錯誤', true);
  }
}

// ── Staff Management ──
function renderStaffList() {
  const container = document.getElementById('topbar-staff-list');
  if (!container) return;
  container.innerHTML = '';
  staffList.forEach((name, idx) => {
    const chip = document.createElement('button');
    chip.className = 'staff-chip' + (name === currentStaff ? ' active' : '');
    chip.innerHTML = `<span class="staff-chip-avatar" style="background:${STAFF_AVATAR_COLORS[idx % STAFF_AVATAR_COLORS.length]}">${escHtml(name[0]||'')}</span><span class="staff-chip-label">${escHtml(name)}</span><span class="staff-chip-del" title="刪除職員">✕</span>`;
    chip.addEventListener('click', function(e) {
      if (e.target.classList.contains('staff-chip-del')) {
        e.stopPropagation();
        openDeleteStaffConfirm(name);
      } else {
        switchStaff(name);
      }
    });
    chip.addEventListener('mouseenter', function() {
      if (name !== currentStaff && !staffDataCache[name]) fetchData(name).then(data => { staffDataCache[name] = data; }).catch(() => {});
    });
    container.appendChild(chip);
  });
}

async function switchStaff(name) {
  if (name === currentStaff) return;
  currentStaff = name;
  localStorage.setItem('ogsm-current-staff', name);
  selectedGoalId = null;
  selectedStrategy = null;
  renderStaffList();
  if (currentTab === 'stats') { renderStats(); return; }
  if (staffDataCache[name]) {
    state = { strategies: [], ...staffDataCache[name] };
    render();
    fetchData().then(data => { staffDataCache[name] = data; if (currentStaff === name) { state = { strategies: [], ...data }; render(); } }).catch(() => {});
  } else { await loadAndRender(); }
}

async function initStaff() {
  try {
    staffList = await fetchStaffList();
    if (!staffList.length) {
      await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ type: 'add_staff', staff_name: 'Riku' }) });
      staffList = ['Riku'];
    }
    if (!staffList.includes(currentStaff)) {
      currentStaff = staffList[0];
      localStorage.setItem('ogsm-current-staff', currentStaff);
    }
  } catch(e) {
    staffList = [currentStaff];
  }
}

function openAddStaffModal() {
  document.getElementById('new-staff-name').value = '';
  openOverlay('modal-add-staff');
}
function closeAddStaffModal() { closeOverlay('modal-add-staff'); }
async function saveNewStaff() {
  const name = document.getElementById('new-staff-name').value.trim();
  if (!name) { showToast('❌ 請填寫職員名稱', true); return; }
  if (staffList.includes(name)) { showToast('❌ 此職員已存在', true); return; }
  const btn = document.getElementById('add-staff-save-btn');
  btn.disabled = true; btn.textContent = '新增中';
  try {
    const res = await postData({ type: 'add_staff', staff_name: name });
    if (res.success) {
      staffList.push(name);
      renderStaffList();
      showToast('✅ 職員新增成功');
      closeAddStaffModal();
    } else { showToast('❌ '+(res.message||'新增失敗'), true); }
  } catch(e) { showToast('❌ 網路錯誤', true); }
  finally { btn.disabled = false; btn.textContent = '新增職員'; }
}

function openDeleteStaffConfirm(name) {
  document.getElementById('confirm-delete-desc').textContent =
    `確定要刪除職員「${name}」？\n此操作將一併刪除該職員的所有資料（試算表），無法復原。`;
  pendingDeleteFn = () => deleteStaff(name);
  openOverlay('modal-confirm-delete');
}

async function deleteStaff(name) {
  try {
    const res = await fetch(GAS_URL, {
      method: 'POST',
      body: JSON.stringify({ type: 'delete_staff', staff_name: name })
    });
    const data = await res.json();
    if (data.success) {
      staffList = staffList.filter(n => n !== name);
      if (currentStaff === name) {
        currentStaff = staffList[0] || '';
        localStorage.setItem('ogsm-current-staff', currentStaff);
        selectedGoalId = null;
        selectedStrategy = null;
      }
      renderStaffList();
      if (currentStaff) {
        await loadAndRender();
      } else {
        document.getElementById('obj-section').style.display = 'none';
        document.getElementById('three-col-wrap').innerHTML = '';
      }
      showToast('✅ 職員已刪除');
    } else {
      showToast('❌ ' + (data.message || '刪除失敗'), true);
    }
  } catch(e) {
    showToast('❌ 網路錯誤', true);
  }
}

// ── Drag-and-Drop Reorder ──
function setupDragDrop(container, onDrop) {
  let dragSrc = null;
  const items = container.querySelectorAll('[data-drag-id]');
  if (items.length < 2) return;
  items.forEach(function(item) {
    item.setAttribute('draggable', 'true');
    item.addEventListener('dragstart', function(e) {
      if (e.target.closest('[contenteditable="true"]')) { e.preventDefault(); return; }
      dragSrc = item;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', item.dataset.dragId);
      requestAnimationFrame(function() { item.classList.add('drag-ghost'); });
    });
    item.addEventListener('dragend', function() {
      item.classList.remove('drag-ghost');
      container.querySelectorAll('.drag-indicator').forEach(function(el) { el.remove(); });
      dragSrc = null;
    });
    item.addEventListener('dragover', function(e) {
      e.preventDefault();
      if (!dragSrc || dragSrc === item) return;
      container.querySelectorAll('.drag-indicator').forEach(function(el) { el.remove(); });
      const rect = item.getBoundingClientRect();
      const ind = document.createElement('div');
      ind.className = 'drag-indicator';
      if (e.clientY < rect.top + rect.height / 2) item.before(ind);
      else item.after(ind);
    });
    item.addEventListener('drop', function(e) {
      e.preventDefault(); e.stopPropagation();
      if (!dragSrc || dragSrc === item) return;
      container.querySelectorAll('.drag-indicator').forEach(function(el) { el.remove(); });
      const allIds = Array.from(container.querySelectorAll('[data-drag-id]')).map(function(el) { return el.dataset.dragId; });
      const srcId = dragSrc.dataset.dragId;
      const dstId = item.dataset.dragId;
      const rect = item.getBoundingClientRect();
      const insertBefore = e.clientY < rect.top + rect.height / 2;
      const ordered = allIds.filter(function(id) { return id !== srcId; });
      const dstPos = ordered.indexOf(dstId);
      ordered.splice(insertBefore ? dstPos : dstPos + 1, 0, srcId);
      onDrop(ordered);
    });
  });
}

function onGoalsDrop(newOrder) {
  const idToGoal = {};
  state.goals.forEach(function(g) { idToGoal[g.id] = g; });
  state.goals = newOrder.map(function(id) { return idToGoal[id]; }).filter(Boolean);
  renderColumns();
  postData({ type: 'reorder_goals', goal_ids: newOrder }).catch(function() { showToast('❌ 排序儲存失敗', true); });
}

function onStrategiesDrop(newOrder) {
  const goalId = selectedGoalId;
  const goalActs = state.actions.filter(function(a) { return a.goal_id === goalId; });
  const otherActs = state.actions.filter(function(a) { return a.goal_id !== goalId; });
  const byStrat = {};
  goalActs.forEach(function(a) {
    const s = a.strategy_name || '（未分類）';
    if (!byStrat[s]) byStrat[s] = [];
    byStrat[s].push(a);
  });
  const reordered = [];
  newOrder.forEach(function(s) { if (byStrat[s]) byStrat[s].forEach(function(a) { reordered.push(a); }); });
  const reorderedIds = new Set(reordered.map(function(a) { return a.id; }));
  goalActs.forEach(function(a) { if (!reorderedIds.has(a.id)) reordered.push(a); });
  state.actions = otherActs.concat(reordered);
  renderColumns();
  postData({ type: 'reorder_strategies', goal_id: goalId, strategy_names: newOrder }).catch(function() { showToast('❌ 排序儲存失敗', true); });
}

function onActionsDrop(newOrder, currentActions) {
  const currentIds = new Set(currentActions.map(function(a) { return a.id; }));
  const idToAction = {};
  currentActions.forEach(function(a) { idToAction[a.id] = a; });
  const otherActs = state.actions.filter(function(a) { return !currentIds.has(a.id); });
  const reordered = newOrder.map(function(id) { return idToAction[id]; }).filter(Boolean);
  state.actions = otherActs.concat(reordered);
  renderColumns();
  postData({ type: 'reorder_actions', action_ids: newOrder }).catch(function() { showToast('❌ 排序儲存失敗', true); });
}

// ── Main ──
async function loadAndRender() {
  try {
    const data = await fetchData();
    staffDataCache[currentStaff] = data;
    state = { strategies: [], ...data };
    render();
  } catch(e) {
    document.getElementById('three-col-wrap').innerHTML =
      `<div class="loading-overlay" style="grid-column:1/-1;color:#f5876e;">⚠️ 載入失敗：${escHtml(e.message)}</div>`;
  }
}

async function init() {
  await initStaff();
  renderStaffList();
  await loadAndRender();
  initOgsmTooltips();
}

init();

function renderMarkdown(text) {
  const lines = text.split('\n');
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    // Table: detect header row followed by separator
    if (i + 1 < lines.length && /^\|.+\|/.test(line) && /^\|[\s\-|:]+\|/.test(lines[i + 1])) {
      const rows = [];
      rows.push(line);
      i += 2; // skip separator
      while (i < lines.length && /^\|.+\|/.test(lines[i])) { rows.push(lines[i++]); }
      const parseRow = (r) => r.replace(/^\||\|$/g, '').split('|').map(c => escHtml(c.trim()));
      const header = parseRow(rows[0]);
      const body = rows.slice(1).map(parseRow);
      let tbl = '<table><thead><tr>' + header.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
      body.forEach(row => { tbl += '<tr>' + row.map(c => `<td>${c}</td>`).join('') + '</tr>'; });
      out.push(tbl + '</tbody></table>');
      continue;
    }
    // Unordered list block
    if (/^- /.test(line)) {
      const items = [];
      while (i < lines.length && /^- /.test(lines[i])) { items.push(escHtml(lines[i].slice(2))); i++; }
      out.push('<ul>' + items.map(it => `<li>${it}</li>`).join('') + '</ul>');
      continue;
    }
    // Headings
    const h2 = line.match(/^## (.+)/);
    if (h2) { out.push(`<h2>${escHtml(h2[1])}</h2>`); i++; continue; }
    const h3 = line.match(/^### (.+)/);
    if (h3) { out.push(`<h3>${escHtml(h3[1])}</h3>`); i++; continue; }
    // HR
    if (/^---+$/.test(line.trim())) { out.push('<hr>'); i++; continue; }
    // Blank line
    if (line.trim() === '') { out.push('<br>'); i++; continue; }
    // Normal line with inline bold
    let s = escHtml(line).replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    out.push(s + '<br>');
    i++;
  }
  return out.join('');
}

// ── Sidebar ──
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

function switchSection(section) {
  const isPersonal = section === 'personal';
  document.getElementById('section-personal').style.display = isPersonal ? '' : 'none';
  document.getElementById('section-department').style.display = isPersonal ? 'none' : 'flex';
  document.getElementById('nav-personal').classList.toggle('active', isPersonal);
  document.getElementById('nav-department').classList.toggle('active', !isPersonal);
  if (!isPersonal) renderMeetingSection();
}

// ── Chat Panel ──
var currentConversationId = null;

function toggleChat() {
  const panel = document.getElementById('chat-panel');
  const btn = document.getElementById('chat-toggle-btn');
  const isOpen = panel.classList.toggle('open');
  btn.classList.toggle('active', isOpen);
  if (!isOpen) {
    panel.style.width = '';
    const inner = panel.querySelector('.chat-panel-inner');
    if (inner) inner.style.width = '';
  }
}

function sendChatMessage() {
  const input = document.getElementById('chat-input');
  const msg = input.value.trim();
  if (!msg) return;

  const messages = document.getElementById('chat-messages');

  const userBubble = document.createElement('div');
  userBubble.className = 'chat-bubble user';
  userBubble.textContent = msg;
  messages.appendChild(userBubble);

  input.value = '';
  input.style.height = '38px';
  messages.scrollTop = messages.scrollHeight;

  const thinking = document.createElement('div');
  thinking.className = 'chat-thinking';
  thinking.innerHTML = '<span></span><span></span><span></span>';
  messages.appendChild(thinking);
  messages.scrollTop = messages.scrollHeight;

  postData({ type: 'ai_chat', message: msg, staff: currentStaff, conversationId: currentConversationId }).then(res => {
    thinking.remove();
    if (res.success && res.conversationId) currentConversationId = res.conversationId;
    const aiBubble = document.createElement('div');
    aiBubble.className = 'chat-bubble ai';
    aiBubble.innerHTML = res.success ? renderMarkdown(res.reply) : ('❌ ' + escHtml(res.error || '發生錯誤'));
    messages.appendChild(aiBubble);
    messages.scrollTop = messages.scrollHeight;
  }).catch(() => {
    thinking.remove();
    const aiBubble = document.createElement('div');
    aiBubble.className = 'chat-bubble ai';
    aiBubble.textContent = '❌ 網路錯誤，請稍後再試';
    messages.appendChild(aiBubble);
    messages.scrollTop = messages.scrollHeight;
  });
}

document.addEventListener('DOMContentLoaded', () => {
  const handle = document.getElementById('chat-resize-handle');
  const panel = document.getElementById('chat-panel');
  if (handle && panel) {
    let startX, startW;
    handle.addEventListener('mousedown', (e) => {
      if (!panel.classList.contains('open')) return;
      startX = e.clientX;
      startW = panel.offsetWidth;
      handle.classList.add('dragging');
      panel.style.transition = 'none';
      const onMove = (e) => {
        const w = Math.min(480, Math.max(420, startW - (e.clientX - startX)));
        panel.style.width = w + 'px';
        panel.querySelector('.chat-panel-inner').style.width = w + 'px';
      };
      const onUp = () => {
        handle.classList.remove('dragging');
        panel.style.transition = '';
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  }

  const chatInput = document.getElementById('chat-input');
  if (!chatInput) return;
  chatInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      sendChatMessage();
    }
  });
  chatInput.addEventListener('input', () => {
    chatInput.style.height = '38px';
    chatInput.style.height = Math.min(chatInput.scrollHeight, 110) + 'px';
  });
});

// ── Meeting Section ──

const MEETING_DEFAULT_ORDER = ['Luka', 'Riku', 'Cathy', 'Yumin'];
const MEETING_STATUS_OPTIONS = ['未開始', '進行中', '待確認解法', '已解決（待觀察）', '已解決（完全改善）', '目前無解'];
let meetingPickerMember = null;
let meetingAddRowMember = null;
let meetingTlEditId = null;

function getMeetingWeekKey() {
  return isoDate(getWeekStart(meetingWeekOffset));
}

function meetingNavWeek(dir) {
  meetingWeekOffset += dir;
  renderMeetingSection();
}

function getMeetingReportData() {
  try { return JSON.parse(localStorage.getItem('meeting-report-v2-' + getMeetingWeekKey()) || '{}'); }
  catch(e) { return {}; }
}

function saveMeetingReportData(data) {
  localStorage.setItem('meeting-report-v2-' + getMeetingWeekKey(), JSON.stringify(data));
}

function getMemberRows(data, memberName) {
  const d = data[memberName];
  if (!d) return [];
  if (d.rows) return d.rows;
  if (d.ogsmItems && d.ogsmItems.length) {
    return d.ogsmItems.map(function(item) {
      const text = item.type === 'M' ? (item.actionName || '') : (item.name || item.text || '');
      return { project: '', task: text, status: d.status || '未開始', bottleneck: d.bottleneck || '' };
    });
  }
  return [];
}

function getMeetingStatusClass(status) {
  const map = {
    '未開始': 'ms-unstart',
    '進行中': 'ms-inprogress',
    '待確認解法': 'ms-pending',
    '已解決（待觀察）': 'ms-resolved-w',
    '已解決（完全改善）': 'ms-resolved-f',
    '目前無解': 'ms-nofix'
  };
  return map[status] || 'ms-unstart';
}

function getMeetingWeekNumber(d) {
  const jan1 = new Date(d.getFullYear(), 0, 1);
  return Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7);
}

function getMeetingOrderedMembers() {
  const all = staffList.length ? [...staffList] : [...MEETING_DEFAULT_ORDER];
  const saved = JSON.parse(localStorage.getItem('meeting-rows-order') || 'null');
  if (!saved) return reorderMeetingByDefault(all);
  const result = saved.filter(function(n) { return all.includes(n); });
  all.forEach(function(n) { if (!result.includes(n)) result.push(n); });
  return result;
}

function reorderMeetingByDefault(members) {
  const result = [];
  MEETING_DEFAULT_ORDER.forEach(function(n) { if (members.includes(n)) result.push(n); });
  members.forEach(function(n) { if (!result.includes(n)) result.push(n); });
  return result;
}

function saveMeetingRowsOrder(order) {
  localStorage.setItem('meeting-rows-order', JSON.stringify(order));
}

function renderMeetingSection() {
  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const weekNum = getMeetingWeekNumber(weekStart);
  const yr = weekStart.getFullYear();

  const labelEl = document.getElementById('meeting-week-label');
  if (labelEl) {
    labelEl.textContent = yr + '年第' + weekNum + '週・' + fmtMD(weekStart) + '（四）～' + fmtMD(weekEnd) + '（三）';
  }
  const awtEl = document.getElementById('meeting-announce-week-title');
  if (awtEl) {
    awtEl.textContent = yr + '年第' + weekNum + '週 佈達事項';
  }

  renderMeetingScore();
  renderMeetingRows();
  renderMeetingTimelineBar();
  renderMeetingAnnounce();
}

function renderMeetingScore() {
  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const startStr = isoDate(weekStart);
  const endStr = isoDate(weekEnd);

  const members = staffList.length ? staffList : MEETING_DEFAULT_ORDER;
  let total = 0;
  members.forEach(function(name) {
    const items = getPersonStats(name).filter(function(i) {
      const d = i.launchDate || i.date;
      return d >= startStr && d <= endStr;
    });
    total += items.reduce(function(s, i) { return s + (i.score || 0); }, 0);
  });

  const el = document.getElementById('meeting-total-score');
  if (el) el.textContent = total;
}

function openDeptScoreModal() {
  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const startStr = isoDate(weekStart);
  const endStr = isoDate(weekEnd);

  const titleEl = document.getElementById('dept-score-modal-title');
  if (titleEl) titleEl.textContent = '本週部門上線項目（' + startStr + ' ~ ' + endStr + '）';

  const members = staffList.length ? staffList : MEETING_DEFAULT_ORDER;
  let bodyHtml = '';
  members.forEach(function(name) {
    const items = getPersonStats(name).filter(function(i) {
      const d = i.launchDate || i.date;
      return d >= startStr && d <= endStr;
    });
    const memberTotal = items.reduce(function(s, i) { return s + (i.score || 0); }, 0);
    const color = avatarColor(name);
    const itemsHtml = items.length
      ? items.map(function(item) {
          const typeClass = item.type && item.type.startsWith('(超大型)') ? ' stats-item-type-xlarge'
            : item.type && item.type.startsWith('(大型)') ? ' stats-item-type-large'
            : item.type && item.type.startsWith('(中型)') ? ' stats-item-type-medium'
            : item.type && item.type.startsWith('(小型)') ? ' stats-item-type-small' : '';
          return '<div class="stats-item-row">' +
            '<div class="stats-item-date">' + escHtml(item.launchDate ? fmtDate(item.launchDate) : (item.date ? fmtDate(item.date) : '')) + '</div>' +
            '<div class="stats-platform-badge">' + escHtml(item.platform || '') + '</div>' +
            (item.target ? '<div class="stats-item-target">' + escHtml(item.target) + '</div>' : '<div class="stats-item-target"></div>') +
            '<div class="stats-item-desc">' + renderDescHtml(item.description || '') + '</div>' +
            '<div class="stats-item-type' + typeClass + '">' + escHtml(item.type || '') + '</div>' +
            '<div class="stats-item-score">+' + (item.score || 0) + '分</div>' +
            '</div>';
        }).join('')
      : '<div class="dept-score-empty">本週無上線項目</div>';
    bodyHtml +=
      '<div class="dept-score-member-section">' +
        '<div class="dept-score-member-header">' +
          '<div class="dept-score-member-avatar" style="background:' + color + '">' + escHtml(initials(name)) + '</div>' +
          '<div class="dept-score-member-name">' + escHtml(name) + '</div>' +
          '<div class="dept-score-member-total">' + memberTotal + ' 分</div>' +
        '</div>' +
        itemsHtml +
      '</div>';
  });

  const bodyEl = document.getElementById('dept-score-modal-body');
  if (bodyEl) bodyEl.innerHTML = bodyHtml;

  const modal = document.getElementById('dept-score-modal');
  if (modal) modal.style.display = 'flex';
}

function closeDeptScoreModal() {
  const modal = document.getElementById('dept-score-modal');
  if (modal) modal.style.display = 'none';
}

async function openDeptNotesModal() {
  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const startStr = isoDate(weekStart);
  const endStr = isoDate(weekEnd);
  const weekRangeStr = startStr + '~' + endStr;

  const titleEl = document.getElementById('dept-notes-modal-title');
  if (titleEl) titleEl.textContent = '本週部門成果/發現問題（' + startStr + ' ~ ' + endStr + '）';

  const modal = document.getElementById('dept-notes-modal');
  if (modal) modal.style.display = 'flex';

  const bodyEl = document.getElementById('dept-notes-modal-body');
  if (bodyEl) bodyEl.innerHTML = '<div style="text-align:center;padding:24px;color:var(--text3)">載入中…</div>';

  const members = staffList.length ? staffList : MEETING_DEFAULT_ORDER;

  await Promise.all(members.map(async function(name) {
    const cacheKey = name + '-' + weekRangeStr;
    if (weekNoteCache[cacheKey] === undefined) {
      try {
        const res = await fetch(GAS_URL + '?api=1&action=get_week_note&staff=' + encodeURIComponent(name) + '&weekStart=' + weekRangeStr + '&_t=' + Date.now(), { cache: 'no-store' });
        const data = await res.json();
        weekNoteCache[cacheKey] = data.content || '';
      } catch(e) {
        weekNoteCache[cacheKey] = '';
      }
    }
  }));

  let filledCount = 0;
  let bodyHtml = '';
  members.forEach(function(name) {
    const cacheKey = name + '-' + weekRangeStr;
    const content = (weekNoteCache[cacheKey] || '').replace(/<a\s/gi, '<a target="_blank" rel="noopener" ');
    if (weekNoteCache[cacheKey]) filledCount++;
    const color = avatarColor(name);
    bodyHtml +=
      '<div class="dept-score-member-section">' +
        '<div class="dept-score-member-header">' +
          '<div class="dept-score-member-avatar" style="background:' + color + '">' + escHtml(initials(name)) + '</div>' +
          '<div class="dept-score-member-name">' + escHtml(name) + '</div>' +
        '</div>' +
        (content
          ? '<div class="dept-notes-content">' + content + '</div>'
          : '<div class="dept-score-empty">本週尚未填寫</div>') +
      '</div>';
  });

  if (bodyEl) bodyEl.innerHTML = bodyHtml;

  const countEl = document.getElementById('meeting-notes-count');
  if (countEl) countEl.textContent = filledCount + '/' + members.length;
}

function closeDeptNotesModal() {
  const modal = document.getElementById('dept-notes-modal');
  if (modal) modal.style.display = 'none';
}

function renderMeetingRows() {
  const container = document.getElementById('meeting-rows');
  if (!container) return;

  const members = getMeetingOrderedMembers();
  const data = getMeetingReportData();

  let html = '<div class="mtable">' +
    '<div class="mtable-head">' +
    '<div class="mth">成員</div>' +
    '<div class="mth">專案</div>' +
    '<div class="mth">本週任務</div>' +
    '<div class="mth">狀態</div>' +
    '<div class="mth">瓶頸 / 備註</div>' +
    '<div class="mth"></div>' +
    '</div>';

  members.forEach(function(name, memberIdx) {
    const rows = getMemberRows(data, name);
    const color = avatarColor(name);

    if (!rows.length) {
      html += '<div class="mtr mtr-first' + (memberIdx === 0 ? ' mtr-first-overall' : '') + '">' +
        '<div class="mtd mtd-member">' +
          '<div class="mrow-avatar" style="background:' + color + '">' + escHtml(name[0] || '') + '</div>' +
          '<div class="mrow-name">' + escHtml(name) + '</div>' +
        '</div>' +
        '<div class="mtd"><span class="mrow-placeholder">尚無任務</span></div>' +
        '<div class="mtd"></div><div class="mtd"></div><div class="mtd"></div>' +
        '<div class="mtd mtd-act">' +
          '<button class="mrow-add-btn" onclick="openMeetingAddRow(' + JSON.stringify(name) + ')" title="新增任務">+</button>' +
        '</div></div>';
      return;
    }

    rows.forEach(function(row, rowIdx) {
      const isFirst = rowIdx === 0;
      const statusCls = getMeetingStatusClass(row.status || '未開始');
      html += '<div class="mtr' + (isFirst ? ' mtr-first' + (memberIdx === 0 ? ' mtr-first-overall' : '') : '') + '">';
      html += '<div class="mtd mtd-member">';
      if (isFirst) {
        html += '<div class="mrow-avatar" style="background:' + color + '">' + escHtml(name[0] || '') + '</div>' +
          '<div class="mrow-name">' + escHtml(name) + '</div>';
      }
      html += '</div>';
      html += '<div class="mtd mtd-project">' +
        '<span class="mrow-editable" contenteditable="true" spellcheck="false" ' +
        'data-field="project" data-member="' + escHtml(name) + '" data-rowidx="' + rowIdx + '" ' +
        'data-placeholder="專案名稱" onblur="saveMeetingRowField(this)" ' +
        'onkeydown="if(event.key===\'Enter\'){event.preventDefault();this.blur()}">' +
        escHtml(row.project || '') + '</span></div>';
      html += '<div class="mtd mtd-task">' +
        '<span class="mrow-editable" contenteditable="true" spellcheck="false" ' +
        'data-field="task" data-member="' + escHtml(name) + '" data-rowidx="' + rowIdx + '" ' +
        'data-placeholder="本週任務描述" onblur="saveMeetingRowField(this)">' +
        escHtml(row.task || '') + '</span></div>';
      html += '<div class="mtd mtd-status">' +
        '<span class="mstatus-badge ' + statusCls + '" ' +
        'data-member="' + escHtml(name) + '" data-rowidx="' + rowIdx + '" ' +
        'onclick="cycleMeetingRowStatus(this)" title="點擊切換狀態">' +
        escHtml(row.status || '未開始') + '</span></div>';
      html += '<div class="mtd mtd-note">' +
        '<span class="mrow-editable" contenteditable="true" spellcheck="false" ' +
        'data-field="bottleneck" data-member="' + escHtml(name) + '" data-rowidx="' + rowIdx + '" ' +
        'data-placeholder="瓶頸 / 備註" onblur="saveMeetingRowField(this)">' +
        escHtml(row.bottleneck || '') + '</span></div>';
      html += '<div class="mtd mtd-act">';
      if (isFirst) {
        html += '<button class="mrow-add-btn" onclick="openMeetingAddRow(' + JSON.stringify(name) + ')" title="新增任務">+</button>';
      }
      html += '<button class="mrow-del-btn" onclick="deleteMeetingRow(' + JSON.stringify(name) + ',' + rowIdx + ')" title="刪除此列">✕</button>';
      html += '</div></div>';
    });
  });

  html += '</div>';
  container.innerHTML = html;
}

function saveMeetingRowField(el) {
  const memberName = el.dataset.member;
  const rowIdx = parseInt(el.dataset.rowidx);
  const field = el.dataset.field;
  const value = el.textContent.trim();
  const data = getMeetingReportData();
  if (!data[memberName]) data[memberName] = { rows: [] };
  if (!data[memberName].rows) data[memberName].rows = getMemberRows(data, memberName);
  if (rowIdx >= 0 && rowIdx < data[memberName].rows.length) {
    data[memberName].rows[rowIdx][field] = value;
    saveMeetingReportData(data);
  }
}

function cycleMeetingRowStatus(el) {
  const memberName = el.dataset.member;
  const rowIdx = parseInt(el.dataset.rowidx);
  const data = getMeetingReportData();
  if (!data[memberName]) return;
  const rows = getMemberRows(data, memberName);
  if (rowIdx < 0 || rowIdx >= rows.length) return;
  const idx = MEETING_STATUS_OPTIONS.indexOf(rows[rowIdx].status || '未開始');
  const next = MEETING_STATUS_OPTIONS[(idx + 1) % MEETING_STATUS_OPTIONS.length];
  rows[rowIdx].status = next;
  data[memberName].rows = rows;
  saveMeetingReportData(data);
  el.textContent = next;
  el.className = 'mstatus-badge ' + getMeetingStatusClass(next);
}

function deleteMeetingRow(memberName, rowIdx) {
  const data = getMeetingReportData();
  if (!data[memberName]) return;
  const rows = getMemberRows(data, memberName);
  rows.splice(rowIdx, 1);
  data[memberName].rows = rows;
  saveMeetingReportData(data);
  renderMeetingRows();
}

function openMeetingAddRow(memberName) {
  meetingAddRowMember = memberName;
  const modal = document.getElementById('meeting-addrow-modal');
  if (!modal) return;
  document.getElementById('meeting-addrow-title').textContent = memberName + ' — 新增任務';
  document.getElementById('addrow-project').value = '';
  document.getElementById('addrow-task').value = '';
  document.getElementById('addrow-status').value = '未開始';
  document.getElementById('addrow-note').value = '';
  modal.style.display = 'flex';
  setTimeout(function() { document.getElementById('addrow-project').focus(); }, 50);
}

function closeMeetingAddRow() {
  const modal = document.getElementById('meeting-addrow-modal');
  if (modal) modal.style.display = 'none';
  meetingAddRowMember = null;
}

function submitMeetingAddRow() {
  if (!meetingAddRowMember) return;
  const project = document.getElementById('addrow-project').value.trim();
  const task = document.getElementById('addrow-task').value.trim();
  const status = document.getElementById('addrow-status').value;
  const bottleneck = document.getElementById('addrow-note').value.trim();
  const data = getMeetingReportData();
  if (!data[meetingAddRowMember]) data[meetingAddRowMember] = { rows: [] };
  if (!data[meetingAddRowMember].rows) data[meetingAddRowMember].rows = getMemberRows(data, meetingAddRowMember);
  data[meetingAddRowMember].rows.push({ project, task, status, bottleneck });
  saveMeetingReportData(data);
  closeMeetingAddRow();
  renderMeetingRows();
  showToast('✅ 任務已新增');
}

function openOgsmPicker(memberName) {
  meetingPickerMember = memberName;
  const modal = document.getElementById('meeting-ogsm-picker');
  const titleEl = document.getElementById('meeting-picker-title');
  const bodyEl = document.getElementById('meeting-picker-body');
  if (!modal) return;
  titleEl.textContent = memberName + ' — 選擇 OGSM 項目';
  const goals = state.goals || [];
  const strategies = state.strategies || [];
  const actions = (state.actions || []).filter(function(a) {
    return !a.assignee || a.assignee === memberName || (a.assignee && a.assignee.includes(memberName));
  });
  let html = '';
  if (goals.length) {
    html += '<div class="picker-section"><div class="picker-section-label"><span class="picker-type-badge badge-g">G</span> 支線目標</div>' +
      goals.map(function(g) {
        return '<div class="picker-item" onclick="addOgsmItemToMember(' + JSON.stringify({ type: 'G', id: g.id, name: g.name }) + ')"><span class="picker-item-text">' + escHtml(g.name || '') + '</span></div>';
      }).join('') + '</div>';
  }
  if (strategies.length) {
    html += '<div class="picker-section"><div class="picker-section-label"><span class="picker-type-badge badge-s">S</span> 策略</div>' +
      strategies.map(function(s) {
        return '<div class="picker-item" onclick="addOgsmItemToMember(' + JSON.stringify({ type: 'S', goalId: s.goal_id, name: s.name }) + ')"><span class="picker-item-text">' + escHtml(s.name || '') + '</span></div>';
      }).join('') + '</div>';
  }
  if (actions.length) {
    html += '<div class="picker-section"><div class="picker-section-label"><span class="picker-type-badge badge-m">M</span> 行動項目</div>' +
      actions.map(function(a) {
        const aName = a.action_name || a.actionName || '';
        return '<div class="picker-item" onclick="addOgsmItemToMember(' + JSON.stringify({ type: 'M', id: a.id, actionName: aName }) + ')"><span class="picker-item-text">' + escHtml(aName) + '</span></div>';
      }).join('') + '</div>';
  }
  if (!html) html = '<div class="picker-empty">尚無 OGSM 資料</div>';
  bodyEl.innerHTML = html;
  modal.style.display = 'flex';
}

function closeOgsmPicker() {
  const modal = document.getElementById('meeting-ogsm-picker');
  if (modal) modal.style.display = 'none';
  meetingPickerMember = null;
}

function addOgsmItemToMember(item) {
  if (!meetingPickerMember) return;
  openMeetingAddRow(meetingPickerMember);
  closeOgsmPicker();
  setTimeout(function() {
    const projectInput = document.getElementById('addrow-project');
    if (projectInput) projectInput.value = item.name || item.actionName || '';
  }, 60);
}

function switchMeetingTab(tab) {
  ['report', 'announce', 'timeline'].forEach(function(t) {
    const panel = document.getElementById('meeting-tab-' + t);
    const btn = document.getElementById('mtab-' + t);
    if (panel) panel.style.display = (t === tab) ? '' : 'none';
    if (btn) btn.classList.toggle('active', t === tab);
  });
  if (tab === 'timeline') renderTimelineTab();
}

// ── Timeline Functions ──

function getTimelineEntries() {
  try { return JSON.parse(localStorage.getItem('meeting-timeline-entries') || '[]'); }
  catch(e) { return []; }
}

function saveTimelineEntries(entries) {
  localStorage.setItem('meeting-timeline-entries', JSON.stringify(entries));
}

function renderTimelineTab() {
  const listEl = document.getElementById('meeting-tl-list');
  if (!listEl) return;
  const entries = getTimelineEntries();
  if (!entries.length) {
    listEl.innerHTML = '<div class="meeting-tl-empty">尚無時程項目<br>點擊「+ 新增時程」建立第一個時程</div>';
    return;
  }
  listEl.innerHTML = entries.map(function(entry) {
    const preview = (entry.content || '').replace(/\n/g, ' ').slice(0, 60) + ((entry.content || '').length > 60 ? '…' : '');
    return '<div class="meeting-tl-entry">' +
      '<div class="meeting-tl-entry-name">' + escHtml(entry.name || '（未命名）') + '</div>' +
      '<div class="meeting-tl-entry-preview">' + escHtml(preview) + '</div>' +
      '<div class="meeting-tl-entry-actions">' +
        '<button class="meeting-tl-entry-btn" onclick="openTimelineModal(' + JSON.stringify(entry.id) + ')">編輯</button>' +
        '<button class="meeting-tl-entry-btn danger" onclick="deleteTimelineEntry(' + JSON.stringify(entry.id) + ')">刪除</button>' +
      '</div></div>';
  }).join('');
}

function renderMeetingTimelineBar() {
  const select = document.getElementById('mtlbar-select');
  if (!select) return;
  const entries = getTimelineEntries();
  const selectedId = localStorage.getItem('meeting-selected-timeline-' + getMeetingWeekKey()) || '';
  select.innerHTML = '<option value="">— 選擇時程 —</option>' +
    entries.map(function(e) {
      return '<option value="' + escHtml(e.id) + '"' + (e.id === selectedId ? ' selected' : '') + '>' + escHtml(e.name || '（未命名）') + '</option>';
    }).join('');
  if (selectedId) {
    const entry = entries.find(function(e) { return e.id === selectedId; });
    if (entry) showTimelinePreview(entry.content);
  }
}

function onTimelineSelectChange(id) {
  if (id) {
    localStorage.setItem('meeting-selected-timeline-' + getMeetingWeekKey(), id);
    const entry = getTimelineEntries().find(function(e) { return e.id === id; });
    if (entry) showTimelinePreview(entry.content);
  } else {
    clearTimelineBar();
  }
}

function showTimelinePreview(content) {
  const preview = document.getElementById('mtlbar-preview');
  const previewContent = document.getElementById('mtlbar-preview-content');
  if (!preview || !previewContent) return;
  previewContent.textContent = content || '';
  preview.style.display = content ? '' : 'none';
}

function clearTimelineBar() {
  localStorage.removeItem('meeting-selected-timeline-' + getMeetingWeekKey());
  const select = document.getElementById('mtlbar-select');
  if (select) select.value = '';
  const preview = document.getElementById('mtlbar-preview');
  if (preview) preview.style.display = 'none';
}

function openTimelineModal(entryId) {
  meetingTlEditId = entryId;
  const modal = document.getElementById('meeting-tl-modal');
  const titleEl = document.getElementById('meeting-tl-modal-title');
  if (!modal) return;
  if (entryId) {
    const entry = getTimelineEntries().find(function(e) { return e.id === entryId; });
    titleEl.textContent = '編輯時程';
    document.getElementById('tl-entry-name').value = entry ? (entry.name || '') : '';
    document.getElementById('tl-entry-content').value = entry ? (entry.content || '') : '';
  } else {
    titleEl.textContent = '新增時程';
    document.getElementById('tl-entry-name').value = '';
    document.getElementById('tl-entry-content').value = '';
  }
  modal.style.display = 'flex';
  setTimeout(function() { document.getElementById('tl-entry-name').focus(); }, 50);
}

function closeMeetingTimelineModal() {
  const modal = document.getElementById('meeting-tl-modal');
  if (modal) modal.style.display = 'none';
  meetingTlEditId = null;
}

function submitTimelineEntry() {
  const name = document.getElementById('tl-entry-name').value.trim();
  const content = document.getElementById('tl-entry-content').value.trim();
  if (!name) { showToast('❗ 請輸入時程名稱', true); return; }
  const entries = getTimelineEntries();
  if (meetingTlEditId) {
    const idx = entries.findIndex(function(e) { return e.id === meetingTlEditId; });
    if (idx >= 0) { entries[idx].name = name; entries[idx].content = content; }
  } else {
    entries.push({ id: Date.now().toString(), name: name, content: content });
  }
  saveTimelineEntries(entries);
  closeMeetingTimelineModal();
  renderTimelineTab();
  renderMeetingTimelineBar();
  showToast('✅ 時程已儲存');
}

function deleteTimelineEntry(id) {
  if (!confirm('確定要刪除此時程嗎？')) return;
  saveTimelineEntries(getTimelineEntries().filter(function(e) { return e.id !== id; }));
  renderTimelineTab();
  renderMeetingTimelineBar();
  showToast('已刪除時程');
}

function renderMeetingAnnounce() {
  const weekKey = getMeetingWeekKey();
  const editor = document.getElementById('meeting-announce-editor');
  if (editor) {
    const saved = localStorage.getItem('meeting-announce-' + weekKey) || '';
    editor.innerHTML = saved;
  }
  renderMeetingAnnounceHistory();
}

function meetingAnnounceCmd(cmd) {
  const editor = document.getElementById('meeting-announce-editor');
  if (!editor) return;
  editor.focus();
  if (cmd === 'link') {
    const url = prompt('請輸入連結網址（含 https://）：');
    if (url) document.execCommand('createLink', false, url);
  } else {
    document.execCommand(cmd, false, null);
  }
}

function meetingAnnounceSave() {
  const editor = document.getElementById('meeting-announce-editor');
  if (!editor) return;
  const weekKey = getMeetingWeekKey();
  localStorage.setItem('meeting-announce-' + weekKey, editor.innerHTML);
  showToast('✅ 佈達事項已儲存');
  renderMeetingAnnounceHistory();
}

function renderMeetingAnnounceHistory() {
  const listEl = document.getElementById('meeting-announce-history-list');
  if (!listEl) return;

  const currentKey = 'meeting-announce-' + getMeetingWeekKey();
  const historyKeys = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith('meeting-announce-') && k !== currentKey) {
      historyKeys.push(k);
    }
  }
  historyKeys.sort().reverse();

  if (!historyKeys.length) {
    listEl.innerHTML = '<div class="announce-history-empty">尚無歷史紀錄</div>';
    return;
  }

  listEl.innerHTML = historyKeys.map(function(k) {
    const dateStr = k.replace('meeting-announce-', '');
    const content = localStorage.getItem(k) || '';
    return '<details class="announce-history-item">' +
      '<summary class="announce-history-summary">' + escHtml(dateStr) + ' 週</summary>' +
      '<div class="announce-history-content">' + content + '</div>' +
      '</details>';
  }).join('');
}
