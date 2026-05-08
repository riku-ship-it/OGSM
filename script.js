const SUPABASE_URL = 'https://bibemmnkwqugooltrrtl.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJpYmVtbW5rd3F1Z29vbHRycnRsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDYwMDk1MDYsImV4cCI6MjA2MTU4NTUwNn0.x6JKdJQa7xMIp1n1TUqbPKYpPHb7dRqWbLVrxuAGM5E';
const sb = () => ({ url: SUPABASE_URL, key: SUPABASE_KEY });

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
  const staff = currentStaff;
  renderStats();
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stats_items?staff=eq.${encodeURIComponent(staff)}&select=*`, { headers: h });
    const data = await res.json();
    const allData = getStatsData();
    allData[staff] = data.map(item => ({ id: item.id, launchDate: item.launch_date, platform: item.platform, target: item.target, description: item.description, type: item.type_name, score: item.score, date: item.launch_date }));
    saveStatsData(allData);
    if (currentStaff === staff) renderStats();
  } catch(e) {}
}

async function postStatsToBackend(payload) {
  const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}`, 'Content-Type': 'application/json' };
  try {
    if (payload.type === 'add_stats_item') {
      await fetch(`${SUPABASE_URL}/rest/v1/stats_items`, { method: 'POST', headers: h, body: JSON.stringify({ id: payload.id, staff: currentStaff, launch_date: payload.launchDate, platform: payload.platform, target: payload.target, description: payload.description, type_name: payload.type_name, score: payload.score }) });
    } else if (payload.type === 'update_stats_item') {
      await fetch(`${SUPABASE_URL}/rest/v1/stats_items?id=eq.${payload.id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ launch_date: payload.launchDate, platform: payload.platform, target: payload.target, description: payload.description, type_name: payload.type_name, score: payload.score }) });
    } else if (payload.type === 'delete_stats_item') {
      await fetch(`${SUPABASE_URL}/rest/v1/stats_items?id=eq.${payload.id}`, { method: 'DELETE', headers: h });
    }
  } catch(e) {}
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
        '<button class="stats-desc-link-btn" onmousedown="event.preventDefault();descLinkCmd(\'ei-desc\')" title="插入連結"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
        '<div class="stats-desc-wrap" style="flex:2">' +
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
        '<button class="stats-desc-link-btn" onmousedown="event.preventDefault();descLinkCmd(\'sf-desc\')" title="插入連結"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
        '<div class="stats-desc-wrap" style="flex:2">' +
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
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/weekly_notes?staff=eq.${encodeURIComponent(person)}&week_start=eq.${encodeURIComponent(weekStartStr)}&select=content`, { headers: h });
    const data = await res.json();
    weekNoteCache[cacheKey] = (data[0] && data[0].content) || '';
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
  const _mws = getWeekStart(meetingWeekOffset);
  const _mwRangeStr = isoDate(_mws) + '~' + isoDate(getWeekEnd(_mws));
  if (weekRangeStr === _mwRangeStr) {
    const _mNoteId = 'mmn-' + currentStaff.replace(/[^a-zA-Z0-9]/g, '_');
    const _mEditor = document.getElementById(_mNoteId);
    if (_mEditor) _mEditor.innerHTML = editor.innerHTML;
  }
  const person = currentStaff;
  clearTimeout(weekNoteTimers[cacheKey]);
  weekNoteTimers[cacheKey] = setTimeout(async function() {
    try {
      await postData({ type: 'save_week_note', staff: person, weekStart: weekRangeStr, content: weekNoteCache[cacheKey] || '' });
      delete weekNoteTimers[cacheKey];
      const s = document.getElementById('stats-note-save-status');
      if (s) { s.textContent = '已儲存'; setTimeout(function() { if (s) s.textContent = ''; }, 2000); }
    } catch(e) { delete weekNoteTimers[cacheKey]; }
  }, 1500);
}

function _linkPopoverOutsideHandler(e) {
  const pop = document.getElementById('link-popover');
  if (pop && !pop.contains(e.target)) _closeLinkPopover();
}
function _closeLinkPopover() {
  const pop = document.getElementById('link-popover');
  if (pop) pop.remove();
  document.removeEventListener('mousedown', _linkPopoverOutsideHandler);
}
function showLinkPopover(editorEl, onApply) {
  const sel = window.getSelection();
  const hasSelection = !!(sel && sel.toString().trim());
  let savedRange = null;
  let rect = null;
  if (sel && sel.rangeCount > 0) {
    savedRange = sel.getRangeAt(0).cloneRange();
    rect = savedRange.getBoundingClientRect();
  }

  _closeLinkPopover();

  const pop = document.createElement('div');
  pop.id = 'link-popover';
  pop.className = 'link-popover';

  let html = '';
  if (!hasSelection) {
    html += '<div class="link-popover-row"><label>顯示文字</label><input id="lp-text" type="text" placeholder="顯示文字" autocomplete="off"></div>';
  }
  html += '<div class="link-popover-row"><label>連結網址</label><input id="lp-url" type="text" placeholder="https://" autocomplete="off"></div>';
  html += '<div class="link-popover-actions"><button class="link-popover-cancel" id="lp-cancel">取消</button><button class="link-popover-confirm" id="lp-confirm">套用</button></div>';
  pop.innerHTML = html;
  document.body.appendChild(pop);

  if (rect && rect.width > 0) {
    let left = rect.left;
    let top = rect.bottom + 6;
    pop.style.left = left + 'px';
    pop.style.top = top + 'px';
    requestAnimationFrame(function() {
      const pw = pop.offsetWidth;
      if (left + pw > window.innerWidth - 10) {
        pop.style.left = Math.max(10, window.innerWidth - pw - 10) + 'px';
      }
      if (top + pop.offsetHeight > window.innerHeight - 10) {
        pop.style.top = Math.max(10, (rect.top - pop.offsetHeight - 6)) + 'px';
      }
    });
  } else if (editorEl) {
    const er = editorEl.getBoundingClientRect();
    pop.style.left = er.left + 'px';
    pop.style.top = (er.top + 40) + 'px';
  }

  const urlInput = document.getElementById('lp-url');
  const textInput = document.getElementById('lp-text');
  (textInput || urlInput).focus();

  function apply() {
    const url = urlInput.value.trim();
    if (!url) { _closeLinkPopover(); return; }
    const displayText = textInput ? textInput.value.trim() : null;
    _closeLinkPopover();
    if (savedRange) {
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(savedRange);
    }
    editorEl.focus();
    onApply(url, displayText, hasSelection);
  }

  document.getElementById('lp-confirm').addEventListener('click', apply);
  document.getElementById('lp-cancel').addEventListener('click', _closeLinkPopover);
  pop.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') { e.preventDefault(); apply(); }
    if (e.key === 'Escape') { _closeLinkPopover(); }
  });
  setTimeout(function() {
    document.addEventListener('mousedown', _linkPopoverOutsideHandler);
  }, 0);
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
    showLinkPopover(editor, function(url, displayText, hasSelection) {
      if (hasSelection) {
        document.execCommand('createLink', false, url);
      } else {
        const text = displayText || url;
        document.execCommand('insertHTML', false, '<a href="' + url + '">' + text + '</a>');
      }
      scheduleWeekNoteSave();
    });
    return;
  } else if (cmd === 'list') {
    document.execCommand('insertUnorderedList', false, null);
  }
  scheduleWeekNoteSave();
}
function descLinkCmd(editorId) {
  const editor = document.getElementById(editorId);
  if (!editor) return;
  editor.focus();
  showLinkPopover(editor, function(url, displayText, hasSelection) {
    if (hasSelection) {
      document.execCommand('createLink', false, url);
      editor.querySelectorAll('a[href="' + url + '"]').forEach(function(a) {
        a.target = '_blank'; a.rel = 'noopener';
      });
    } else {
      const text = displayText || url;
      document.execCommand('insertHTML', false, '<a href="' + url + '" target="_blank" rel="noopener">' + text + '</a>');
    }
  });
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
  const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
  const targetStaff = staff || currentStaff;

  // 先取得該職員的所有 goal_id
  const goalRes = await fetch(
    `${SUPABASE_URL}/rest/v1/goals?staff=eq.${encodeURIComponent(targetStaff)}&select=id,objective_id,staff,name,color,progress,traffic_light,deadline,sort_order&order=sort_order.asc`,
    { headers: h }
  );
  const goals = await goalRes.json();

  // 取得 objectives
  const objRes = await fetch(
    `${SUPABASE_URL}/rest/v1/objectives?staff=eq.${encodeURIComponent(targetStaff)}&select=*`,
    { headers: h }
  );
  const objectives = await objRes.json();

  // 取得 strategies
  const stratRes = await fetch(
    `${SUPABASE_URL}/rest/v1/strategies?select=*`,
    { headers: h }
  );
  const strategies = await stratRes.json();

  // 用 goal_id 清單取 actions
  let actions = [];
  if (goals && goals.length > 0) {
    const goalIds = goals.map(g => g.id).join(',');
    const actRes = await fetch(
      `${SUPABASE_URL}/rest/v1/actions?goal_id=in.(${goalIds})&select=*&order=sort_order.asc`,
      { headers: h }
    );
    actions = await actRes.json();
  }

  return {
    objectives: Array.isArray(objectives) ? objectives : [],
    goals: Array.isArray(goals) ? goals : [],
    strategies: Array.isArray(strategies) ? strategies : [],
    actions: Array.isArray(actions) ? actions : [],
  };
}
async function fetchStaffList() {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/staff?select=name`, {
    headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` }
  });
  const data = await res.json();
  return data.map(r => r.name);
}
async function postData(payload) {
  const h = {
    'apikey': SUPABASE_KEY,
    'Authorization': `Bearer ${SUPABASE_KEY}`,
    'Content-Type': 'application/json',
    'Prefer': 'return=representation'
  };
  const type = payload.type;

  if (type === 'rename_objective') {
    await fetch(`${SUPABASE_URL}/rest/v1/objectives?id=eq.${payload.obj_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ title: payload.new_title }) });
    return { success: true };
  }
  if (type === 'create_objective') {
    const id = 'Obj_' + Date.now();
    await fetch(`${SUPABASE_URL}/rest/v1/objectives`, { method: 'POST', headers: h, body: JSON.stringify({ id, staff: currentStaff, title: payload.new_title }) });
    return { success: true, obj_id: id };
  }
  if (type === 'add_goal') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals`, { method: 'POST', headers: h, body: JSON.stringify({ id: payload.goal_id, objective_id: payload.obj_id, staff: currentStaff, name: payload.goal_name, color: payload.goal_color, progress: payload.goal_progress || 0, deadline: payload.goal_deadline || null }) });
    return { success: true };
  }
  if (type === 'rename_goal') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${payload.goal_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ name: payload.new_name }) });
    return { success: true };
  }
  if (type === 'delete_goal') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions?goal_id=eq.${payload.goal_id}`, { method: 'DELETE', headers: h });
    await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${payload.goal_id}`, { method: 'DELETE', headers: h });
    await fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${payload.goal_id}`, { method: 'DELETE', headers: h });
    return { success: true };
  }
  if (type === 'update_goal_color') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${payload.goal_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ color: payload.color }) });
    return { success: true };
  }
  if (type === 'update_goal_deadline') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${payload.goal_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ deadline: payload.deadline || null }) });
    return { success: true };
  }
  if (type === 'update_goal_traffic') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${payload.goal_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ traffic_light: payload.traffic_light }) });
    return { success: true };
  }
  if (type === 'add_action') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions`, { method: 'POST', headers: h, body: JSON.stringify({ id: payload.action_id, goal_id: payload.goal_id, strategy_name: payload.strategy_name, action_name: payload.action_name, assignee: payload.assignee, due_date: payload.due_date || null, notes: payload.notes, status: payload.status || '未開始' }) });
    return { success: true };
  }
  if (type === 'update_action') {
    const updates = {};
    if (payload.action_name !== undefined) updates.action_name = payload.action_name;
    if (payload.strategy_name !== undefined) updates.strategy_name = payload.strategy_name;
    if (payload.assignee !== undefined) updates.assignee = payload.assignee;
    if (payload.due_date !== undefined) updates.due_date = payload.due_date || null;
    if (payload.notes !== undefined) updates.notes = payload.notes;
    if (payload.status !== undefined) updates.status = payload.status;
    if (payload.success_def !== undefined) updates.success_def = payload.success_def;
    await fetch(`${SUPABASE_URL}/rest/v1/actions?id=eq.${payload.id}`, { method: 'PATCH', headers: h, body: JSON.stringify(updates) });
    return { success: true };
  }
  if (type === 'rename_action') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions?id=eq.${payload.action_id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ action_name: payload.new_name }) });
    return { success: true };
  }
  if (type === 'delete_action') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions?id=eq.${payload.action_id}`, { method: 'DELETE', headers: h });
    return { success: true };
  }
  if (type === 'rename_strategy') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions?goal_id=eq.${payload.goal_id}&strategy_name=eq.${encodeURIComponent(payload.old_name)}`, { method: 'PATCH', headers: h, body: JSON.stringify({ strategy_name: payload.new_name }) });
    await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${payload.goal_id}&name=eq.${encodeURIComponent(payload.old_name)}`, { method: 'PATCH', headers: h, body: JSON.stringify({ name: payload.new_name }) });
    return { success: true };
  }
  if (type === 'delete_strategy') {
    await fetch(`${SUPABASE_URL}/rest/v1/actions?goal_id=eq.${payload.goal_id}&strategy_name=eq.${encodeURIComponent(payload.strategy_name)}`, { method: 'DELETE', headers: h });
    await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${payload.goal_id}&name=eq.${encodeURIComponent(payload.strategy_name)}`, { method: 'DELETE', headers: h });
    return { success: true };
  }
  if (type === 'update_strategy_status') {
    await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${payload.goal_id}&name=eq.${encodeURIComponent(payload.strategy_name)}`, { method: 'PATCH', headers: h, body: JSON.stringify({ status: payload.status }) });
    return { success: true };
  }
  if (type === 'update_strategy_success_def') {
    await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${payload.goal_id}&name=eq.${encodeURIComponent(payload.strategy_name)}`, { method: 'PATCH', headers: h, body: JSON.stringify({ success_def: payload.success_def }) });
    return { success: true };
  }
  if (type === 'add_staff') {
    await fetch(`${SUPABASE_URL}/rest/v1/staff`, { method: 'POST', headers: h, body: JSON.stringify({ name: payload.staff_name }) });
    return { success: true };
  }
  if (type === 'delete_staff') {
    await fetch(`${SUPABASE_URL}/rest/v1/goals?staff=eq.${encodeURIComponent(payload.staff_name)}&select=id`, { headers: h })
      .then(r => r.json())
      .then(async goals => {
        for (const g of goals) {
          await fetch(`${SUPABASE_URL}/rest/v1/actions?goal_id=eq.${g.id}`, { method: 'DELETE', headers: h });
          await fetch(`${SUPABASE_URL}/rest/v1/strategies?goal_id=eq.${g.id}`, { method: 'DELETE', headers: h });
        }
      });
    await fetch(`${SUPABASE_URL}/rest/v1/goals?staff=eq.${encodeURIComponent(payload.staff_name)}`, { method: 'DELETE', headers: h });
    await fetch(`${SUPABASE_URL}/rest/v1/objectives?staff=eq.${encodeURIComponent(payload.staff_name)}`, { method: 'DELETE', headers: h });
    await fetch(`${SUPABASE_URL}/rest/v1/staff?name=eq.${encodeURIComponent(payload.staff_name)}`, { method: 'DELETE', headers: h });
    return { success: true };
  }
  if (type === 'reorder_goals') {
    await Promise.all(payload.goal_ids.map((id, idx) =>
      fetch(`${SUPABASE_URL}/rest/v1/goals?id=eq.${id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ sort_order: idx }) })
    ));
    return { success: true };
  }
  if (type === 'reorder_actions') {
    await Promise.all(payload.action_ids.map((id, idx) =>
      fetch(`${SUPABASE_URL}/rest/v1/actions?id=eq.${id}`, { method: 'PATCH', headers: h, body: JSON.stringify({ sort_order: idx }) })
    ));
    return { success: true };
  }
  if (type === 'save_week_note') {
    await fetch(`${SUPABASE_URL}/rest/v1/weekly_notes`, { method: 'POST', headers: { ...h, 'Prefer': 'resolution=merge-duplicates' }, body: JSON.stringify({ staff: payload.staff, week_start: payload.weekStart, content: payload.content }) });
    return { success: true };
  }
  if (type === 'save_meeting_note') {
    await fetch(`${SUPABASE_URL}/rest/v1/meeting_notes`, { method: 'POST', headers: h, body: JSON.stringify({ note_type: payload.noteType, week_key: payload.weekKey, member: payload.member || '', content: payload.content }) });
    return { success: true };
  }
  if (type === 'save_meeting_selections') {
    await fetch(`${SUPABASE_URL}/rest/v1/meeting_selections`, { method: 'POST', headers: { ...h, 'Prefer': 'resolution=merge-duplicates' }, body: JSON.stringify({ week_key: payload.weekKey, member: payload.member, selected_action_ids: payload.selectedActionIds, selected_strategy_keys: payload.selectedStrategyKeys }) });
    return { success: true };
  }
  if (type === 'save_meeting_report') {
    return { success: true };
  }
  if (type === 'ai_chat' || type === 'ai_meeting_summary' || type === 'ai_generate_meeting') {
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbywtjAiFsqzEzKiQ4soH7LdRHWiViQTCrte3fL2ySS49nPb_w3Zk7ctX1dyb1A0zDmMXw/exec';
    const res = await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ ...payload, staff: currentStaff }) });
    return await res.json();
  }
  return { success: false, message: 'unknown type: ' + type };
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
  nameEl.onkeydown = function(e) {
    if (e.key==='Enter' && !e.isComposing) { e.preventDefault(); nameEl.blur(); }
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
      nameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !e.isComposing) { e.preventDefault(); nameEl.blur(); }
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
      sNameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !e.isComposing) { e.preventDefault(); sNameEl.blur(); }
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
  const mCol = makeColumn('M', 'Action-行動計劃', 'col-tag-m', mCount);
  const mBody = mCol.querySelector('.col-body');

  if (!selectedGoalId) {
    mBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">←</div><span>請先選擇左側目標</span></div>';
  } else if (!mActions.length) {
    mBody.innerHTML = '<div class="col-empty"><div class="col-empty-icon">📝</div><span>尚無Action-行動計劃</span></div>';
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
        ${a.notes ? `<div class="action-item-notes"><svg class="action-item-notes-icon" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="10" cy="10" r="9" stroke="currentColor" stroke-width="1.5"/><line x1="10" y1="9" x2="10" y2="14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="10" cy="6.5" r="0.75" fill="currentColor"/></svg><span>${escHtml(a.notes)}</span></div>` : ''}
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
      aNameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter' && !e.isComposing) { e.preventDefault(); aNameEl.blur(); }
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
        <span class="ogsm-tooltip-label">Action-行動計劃</span>
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
  document.getElementById('edit-notes').value = a.notes || '';
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
    notes:  document.getElementById('edit-notes').value.trim(),
    status: document.getElementById('edit-status').value,
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
        a.notes         = payload.notes;
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
  document.getElementById('new-action-notes').value = '';
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
    notes:document.getElementById('new-action-notes').value.trim(),
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
  if (currentTab === 'stats') { loadStats(); return; }
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
      await postData({ type: 'add_staff', staff_name: 'Riku' });
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
    const data = await postData({ type: 'delete_staff', staff_name: name });
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
  staffList
    .filter(function(name) { return name !== currentStaff && !staffDataCache[name]; })
    .forEach(function(name) { fetchData(name).then(function(data) { staffDataCache[name] = data; }).catch(function() {}); });
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

var _meetingSyncTimer = null;
function switchSection(section) {
  const isPersonal = section === 'personal';
  document.getElementById('section-personal').style.display = isPersonal ? '' : 'none';
  document.getElementById('section-department').style.display = isPersonal ? 'none' : 'flex';
  document.getElementById('nav-personal').classList.toggle('active', isPersonal);
  document.getElementById('nav-department').classList.toggle('active', !isPersonal);
  if (!isPersonal) {
    renderMeetingSection();
    renderAiSummaryHistory();
    if (!_meetingSyncTimer) _meetingSyncTimer = setInterval(_syncMeetingSelectionsFromServer, 10000);
  } else {
    clearInterval(_meetingSyncTimer); _meetingSyncTimer = null;
  }
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

const MEETING_DEFAULT_ORDER = ['Luka', 'Riku', 'Cathy'];
const MEETING_STATUS_OPTIONS = ['未開始', '進行中', '待確認解法', '已解決（待觀察）', '已解決（完全改善）', '目前無解'];
let meetingPickerMember = null;
let meetingAddRowMember = null;
let meetingTlEditId = null;
let meetingReportCache = null;
let meetingStatusFilter = null;
let meetingCollapsedMembers = {};
let meetingSelectionsCache = {};
let _pendingPushMembers = new Set();
let aiMeetingTempItems = [];
let meetingPickerChecked = { actions: new Set(), strategies: new Set() };
let pickerActiveGoalId = null;
let pickerActiveStrategyKey = null;
let meetingMemberNotesCache = {};
let meetingMemberNoteTimers = {};

function getMeetingWeekKey() {
  return isoDate(getWeekStart(meetingWeekOffset));
}

function meetingNavWeek(dir) {
  meetingWeekOffset += dir;
  meetingReportCache = null;
  renderMeetingSection();
}

function getMeetingReportData() {
  if (meetingReportCache !== null) return meetingReportCache;
  try { return JSON.parse(localStorage.getItem('meeting-report-v2-' + getMeetingWeekKey()) || '{}'); }
  catch(e) { return {}; }
}

function saveMeetingReportData(data) {
  meetingReportCache = data;
  localStorage.setItem('meeting-report-v2-' + getMeetingWeekKey(), JSON.stringify(data));
  postData({ type: 'save_meeting_report', weekKey: getMeetingWeekKey(), data: data }).catch(function() {});
}

async function loadMeetingReportFromBackend() {
  try {
    const stored = localStorage.getItem('meeting-report-v2-' + getMeetingWeekKey());
    meetingReportCache = stored ? JSON.parse(stored) : null;
  } catch(e) {
    meetingReportCache = null;
  }
}

function _pushMemberSelectionsToServer(memberName) {
  const weekKey = getMeetingWeekKey();
  const cache = ((meetingSelectionsCache[weekKey] || {})[memberName]) || {};
  const payload = { type: 'save_meeting_selections', weekKey: weekKey, member: memberName, selectedActionIds: cache.selectedActionIds || [], selectedStrategyKeys: cache.selectedStrategyKeys || [] };
  _pendingPushMembers.add(memberName);
  return postData(payload).finally(function() { _pendingPushMembers.delete(memberName); });
}

async function loadMeetingSelectionsFromServer() {
  const weekKey = getMeetingWeekKey();
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/meeting_selections?week_key=eq.${encodeURIComponent(weekKey)}&select=*`, { headers: h });
    const data = await res.json();
    const selections = {};
    data.forEach(function(row) {
      selections[row.member] = { selectedActionIds: row.selected_action_ids || [], selectedStrategyKeys: row.selected_strategy_keys || [] };
    });
    meetingSelectionsCache[weekKey] = selections;
    try { localStorage.setItem('meeting-selections-v1-' + weekKey, JSON.stringify(meetingSelectionsCache[weekKey])); } catch(e) {}
  } catch(e) {
    if (!meetingSelectionsCache[weekKey]) meetingSelectionsCache[weekKey] = {};
  }
}

async function loadMeetingNotesFromBackend() {
  const weekKey = getMeetingWeekKey();
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/meeting_notes?week_key=eq.${encodeURIComponent(weekKey)}&select=*`, { headers: h });
    const data = await res.json();
    if (!meetingMemberNotesCache[weekKey]) meetingMemberNotesCache[weekKey] = {};
    const announceRow = data.find(function(r) { return r.note_type === 'announce'; });
    if (announceRow !== undefined) {
      meetingMemberNotesCache[weekKey]._announce = announceRow ? announceRow.content : '';
      if (announceRow && announceRow.content) localStorage.setItem('meeting-announce-' + weekKey, announceRow.content);
    }
  } catch(e) {}
}

async function _syncMeetingSelectionsFromServer() {
  const weekKey = getMeetingWeekKey();
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/meeting_selections?week_key=eq.${encodeURIComponent(weekKey)}&select=*`, { headers: h });
    const data = await res.json();
    const serverSelections = {};
    data.forEach(function(row) {
      serverSelections[row.member] = { selectedActionIds: row.selected_action_ids || [], selectedStrategyKeys: row.selected_strategy_keys || [] };
    });
    const prev = JSON.stringify(meetingSelectionsCache[weekKey] || {});
    if (!meetingSelectionsCache[weekKey]) meetingSelectionsCache[weekKey] = {};
    Object.keys(serverSelections).forEach(function(member) {
      if (!_pendingPushMembers.has(member)) {
        meetingSelectionsCache[weekKey][member] = serverSelections[member];
      }
    });
    if (JSON.stringify(meetingSelectionsCache[weekKey]) !== prev) renderMeetingRows();
  } catch(e) { /* silent on periodic sync */ }
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

async function renderMeetingSection() {
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

  const members = getMeetingOrderedMembers();
  const weekKey = getMeetingWeekKey();

  // Pre-populate caches from localStorage for instant render
  if (!meetingSelectionsCache[weekKey]) {
    try { meetingSelectionsCache[weekKey] = JSON.parse(localStorage.getItem('meeting-selections-v1-' + weekKey) || 'null') || {}; }
    catch(e) { meetingSelectionsCache[weekKey] = {}; }
  }
  members.forEach(function(name) {
    if (!staffDataCache[name]) {
      try {
        const cached = localStorage.getItem('staffdata-v1-' + name);
        if (cached) staffDataCache[name] = JSON.parse(cached);
      } catch(e) {}
    }
  });

  // Render immediately with cached data
  renderMeetingScore();
  renderMeetingStatusFilters();
  renderMeetingAnnounce();
  renderMeetingRows();

  // Fetch fresh data in background, then re-render
  const cachePromises = members
    .filter(function(name) { return name !== currentStaff && !staffDataCache[name]; })
    .map(function(name) {
      return fetchData(name).then(function(data) {
        staffDataCache[name] = data;
        try { localStorage.setItem('staffdata-v1-' + name, JSON.stringify(data)); } catch(e) {}
      }).catch(function() {});
    });

  // Fetch stats for all members so dept score total is complete without clicking avatars
  const statsPromises = members.map(function(name) {
    const _sh = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    return fetch(`${SUPABASE_URL}/rest/v1/stats_items?staff=eq.${encodeURIComponent(name)}&select=*`, { headers: _sh })
      .then(function(res) { return res.json(); })
      .then(function(data) {
        const allData = getStatsData();
        allData[name] = data.map(function(item) {
          return { id: item.id, launchDate: item.launch_date, platform: item.platform, target: item.target, description: item.description, type: item.type_name, score: item.score, date: item.launch_date };
        });
        saveStatsData(allData);
      }).catch(function() {});
  });

  // Fetch week notes for all members so dept notes count shows without opening modal
  const weekRangeStr = isoDate(weekStart) + '~' + isoDate(weekEnd);
  const notePromises = members.map(function(name) {
    const cacheKey = name + '-' + weekRangeStr;
    if (weekNoteCache[cacheKey] !== undefined) return Promise.resolve();
    const _nh = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    return fetch(`${SUPABASE_URL}/rest/v1/weekly_notes?staff=eq.${encodeURIComponent(name)}&week_start=eq.${encodeURIComponent(weekRangeStr)}&select=content`, { headers: _nh })
      .then(function(res) { return res.json(); })
      .then(function(data) { weekNoteCache[cacheKey] = (data[0] && data[0].content) || ''; })
      .catch(function() { weekNoteCache[cacheKey] = ''; });
  });

  await Promise.all([loadMeetingReportFromBackend(), loadMeetingNotesFromBackend(), loadMeetingSelectionsFromServer(), ...cachePromises, ...statsPromises, ...notePromises]);
  renderMeetingScore();
  renderMeetingNotesCount();
  renderMeetingStatusFilters();
  renderMeetingAnnounce();
  renderMeetingRows();
}

function renderMeetingNotesCount() {
  const members = staffList.length ? staffList : MEETING_DEFAULT_ORDER;
  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const weekRangeStr = isoDate(weekStart) + '~' + isoDate(weekEnd);
  let filledCount = 0;
  members.forEach(function(name) {
    if (weekNoteCache[name + '-' + weekRangeStr]) filledCount++;
  });
  const countEl = document.getElementById('meeting-notes-count');
  if (countEl) countEl.textContent = filledCount + '/' + members.length;
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

  // Compute last week's score for delta indicator
  const prevWeekStart = getWeekStart(meetingWeekOffset - 1);
  const prevWeekEnd = getWeekEnd(prevWeekStart);
  const prevStartStr = isoDate(prevWeekStart);
  const prevEndStr = isoDate(prevWeekEnd);
  let prevTotal = 0;
  members.forEach(function(name) {
    const items = getPersonStats(name).filter(function(i) {
      const d = i.launchDate || i.date;
      return d >= prevStartStr && d <= prevEndStr;
    });
    prevTotal += items.reduce(function(s, i) { return s + (i.score || 0); }, 0);
  });
  const deltaEl = document.getElementById('meeting-score-delta');
  if (deltaEl) {
    const diff = total - prevTotal;
    if (diff > 0) {
      deltaEl.textContent = '▲ ' + diff;
      deltaEl.className = 'meeting-score-delta up';
    } else if (diff < 0) {
      deltaEl.textContent = '▼ ' + Math.abs(diff);
      deltaEl.className = 'meeting-score-delta down';
    } else {
      deltaEl.textContent = '— 與上週相同';
      deltaEl.className = 'meeting-score-delta flat';
    }
  }
}

function getDeptScoreForWeekOffset(offset) {
  const weekStart = getWeekStart(offset);
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
  return { score: total, startStr: startStr, weekStart: weekStart };
}

function renderDeptScoreChart() {
  const wrap = document.getElementById('dept-score-chart-wrap');
  if (!wrap) return;
  const weeks = [];
  for (let i = 5; i >= 0; i--) {
    weeks.push(getDeptScoreForWeekOffset(meetingWeekOffset - i));
  }
  const scores = weeks.map(function(w) { return w.score; });
  const rawMax = Math.max.apply(null, scores) || 1;
  const gridSteps = 4;
  const maxScore = Math.ceil(rawMax / gridSteps) * gridSteps || gridSteps;
  const W = 280, H = 185, PL = 42, PR = 12, PT = 20, PB = 34;
  const chartW = W - PL - PR;
  const chartH = H - PT - PB;
  const n = weeks.length;
  function px(i) { return PL + (i / (n - 1)) * chartW; }
  function py(v) { return PT + chartH - (v / maxScore) * chartH; }

  let gridHtml = '';
  for (let g = 0; g <= gridSteps; g++) {
    const val = Math.round((g / gridSteps) * maxScore);
    const y = py(val);
    gridHtml += '<line class="dept-score-chart-grid" x1="' + PL + '" y1="' + y + '" x2="' + (W - PR) + '" y2="' + y + '"/>';
    gridHtml += '<text class="dept-score-chart-y-label" x="' + (PL - 5) + '" y="' + (y + 3.5) + '" text-anchor="end">' + val + '</text>';
  }
  const borderHtml = '<rect class="dept-score-chart-border" x="' + PL + '" y="' + PT + '" width="' + chartW + '" height="' + chartH + '"/>';

  const points = weeks.map(function(w, i) { return px(i) + ',' + py(w.score); }).join(' ');
  const areaPoints = 'M' + px(0) + ',' + (PT + chartH) + ' ' +
    weeks.map(function(w, i) { return 'L' + px(i) + ',' + py(w.score); }).join(' ') +
    ' L' + px(n - 1) + ',' + (PT + chartH) + ' Z';
  let dotsHtml = '';
  let labelsHtml = '';
  weeks.forEach(function(w, i) {
    const x = px(i), y = py(w.score);
    dotsHtml += '<circle class="dept-score-chart-dot" cx="' + x + '" cy="' + y + '" r="4"/>';
    const d = w.weekStart;
    const label = (d.getMonth() + 1) + '/' + d.getDate();
    labelsHtml += '<text class="dept-score-chart-label" x="' + x + '" y="' + (H - 8) + '" text-anchor="middle">' + label + '</text>';
    labelsHtml += '<text class="dept-score-chart-value" x="' + x + '" y="' + (y - 8) + '" text-anchor="middle">' + w.score + '</text>';
  });
  wrap.innerHTML = '<svg width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H + '">' +
    '<defs><linearGradient id="scoreChartGrad" x1="0" y1="0" x2="0" y2="1">' +
    '<stop offset="0%" stop-color="var(--blue)" stop-opacity="0.45"/>' +
    '<stop offset="100%" stop-color="var(--blue)" stop-opacity="0"/>' +
    '</linearGradient></defs>' +
    gridHtml + borderHtml +
    '<path class="dept-score-chart-area" d="' + areaPoints + '"/>' +
    '<polyline class="dept-score-chart-line" points="' + points + '"/>' +
    dotsHtml + labelsHtml +
    '</svg>';
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

  renderDeptScoreChart();
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
        const _dh = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
        const res = await fetch(`${SUPABASE_URL}/rest/v1/weekly_notes?staff=eq.${encodeURIComponent(name)}&week_start=eq.${encodeURIComponent(weekRangeStr)}&select=content`, { headers: _dh });
        const data = await res.json();
        weekNoteCache[cacheKey] = (data[0] && data[0].content) || '';
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

function closeAiSummaryModal() {
  const modal = document.getElementById('ai-summary-modal');
  if (modal) modal.style.display = 'none';
}

function renderAiSummaryMarkdown(text) {
  return text
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/^[-*] (.+)$/gm, '<li>$1</li>')
    .replace(/(<li>[\s\S]*?<\/li>)/g, '<ul>$1</ul>')
    .replace(/<\/ul>\s*<ul>/g, '')
    .replace(/\n/g, '<br>');
}

async function renderAiSummaryHistory() {
  const section = document.getElementById('sidebar-history-section');
  const listEl = document.getElementById('sidebar-history-list');
  if (!section || !listEl) return;
  const currentWeekKey = getMeetingWeekKey();
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/meeting_notes?note_type=eq.ai_summary&week_key=neq.${encodeURIComponent(currentWeekKey)}&select=week_key,content&order=week_key.desc`, { headers: h });
    const data = await res.json();
    const items = data.filter(function(i) { return !!i.content; }).map(function(i) { return { weekKey: i.week_key, content: i.content }; });
    if (!items.length) { section.style.display = 'none'; return; }
    section.style.display = '';
    listEl.innerHTML = items.map(function(item, idx) {
      return '<div class="sidebar-history-item" data-idx="' + idx + '">' + escHtml(item.weekKey) + ' 週</div>';
    }).join('');
    listEl.querySelectorAll('.sidebar-history-item').forEach(function(el) {
      const idx = parseInt(el.dataset.idx, 10);
      el.addEventListener('click', function() {
        openHistorySummaryModal(items[idx].weekKey, items[idx].content);
      });
    });
  } catch(e) { section.style.display = 'none'; }
}

function openHistorySummaryModal(weekKey, content) {
  const modal = document.getElementById('ai-summary-modal');
  const bodyEl = document.getElementById('ai-summary-modal-body');
  const titleEl = document.getElementById('ai-summary-modal-title');
  if (!modal || !bodyEl) return;
  if (titleEl) titleEl.textContent = '部門週報摘要（' + weekKey + ' 週）';
  bodyEl.innerHTML = '<div class="ai-summary-content">' + renderAiSummaryMarkdown(content) + '</div>';
  modal.style.display = 'flex';
}

async function generateMeetingSummary() {
  const modal = document.getElementById('ai-summary-modal');
  const bodyEl = document.getElementById('ai-summary-modal-body');
  const titleEl = document.getElementById('ai-summary-modal-title');
  if (!modal || !bodyEl) return;

  const weekStart = getWeekStart(meetingWeekOffset);
  const weekEnd = getWeekEnd(weekStart);
  const startStr = isoDate(weekStart);
  const endStr = isoDate(weekEnd);
  const weekRangeStr = startStr + '~' + endStr;
  const weekKey = getMeetingWeekKey();

  if (titleEl) titleEl.textContent = '部門週報 AI 摘要（' + startStr + ' ~ ' + endStr + '）';
  bodyEl.innerHTML = '<div class="ai-summary-loading">AI 分析中，請稍候…</div>';
  modal.style.display = 'flex';

  const members = getMeetingOrderedMembers();

  // ensure week notes are cached
  await Promise.all(members.map(async function(name) {
    const cacheKey = name + '-' + weekRangeStr;
    if (weekNoteCache[cacheKey] === undefined) {
      try {
        const _gh = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
        const res = await fetch(`${SUPABASE_URL}/rest/v1/weekly_notes?staff=eq.${encodeURIComponent(name)}&week_start=eq.${encodeURIComponent(weekRangeStr)}&select=content`, { headers: _gh });
        const data = await res.json();
        weekNoteCache[cacheKey] = (data[0] && data[0].content) || '';
      } catch(e) { weekNoteCache[cacheKey] = ''; }
    }
  }));

  // compute status counts across all members
  const statusCounts = { '未開始': 0, '進行中': 0, '卡關': 0, '完成': 0 };
  members.forEach(function(name) {
    const data = name === currentStaff ? state : (staffDataCache[name] || {});
    (data.actions || []).forEach(function(a) {
      if (a.action_name && statusCounts[a.status] !== undefined) statusCounts[a.status]++;
    });
  });

  // build per-member data
  const membersData = members.map(function(name) {
    const data = name === currentStaff ? state : (staffDataCache[name] || {});
    const allActions = (data.actions || []).filter(function(a) { return !!a.action_name; });
    const selectedIds = getSelectedActionIds(name);
    const selectedActions = selectedIds.map(function(id) {
      return allActions.find(function(a) { return String(a.id) === String(id); });
    }).filter(Boolean).map(function(a) {
      return { action_name: a.action_name, status: a.status || '未開始', assignee: a.assignee || '' };
    });

    const rawMemberNote = ((meetingMemberNotesCache[weekKey] || {})[name] || '').replace(/<[^>]*>/g, '').trim();
    const rawWeekNote = (weekNoteCache[name + '-' + weekRangeStr] || '').replace(/<[^>]*>/g, '').trim();

    return { name: name, selectedActions: selectedActions, memberNote: rawMemberNote, weekNote: rawWeekNote };
  });

  try {
    const json = await postData({
      type: 'ai_meeting_summary',
      weekRange: startStr + ' ~ ' + endStr,
      members: membersData,
      statusCounts: statusCounts
    });
    if (json.success && json.summary) {
      bodyEl.innerHTML = '<div class="ai-summary-content">' + renderAiSummaryMarkdown(json.summary) + '</div>';
      renderAiSummaryHistory();
    } else {
      bodyEl.innerHTML = '<div class="ai-summary-error">摘要產生失敗：' + escHtml(json.error || '未知錯誤') + '</div>';
    }
  } catch(e) {
    bodyEl.innerHTML = '<div class="ai-summary-error">連線失敗，請稍後再試</div>';
  }
}

function getSelectedActionIds(memberName) {
  const weekKey = getMeetingWeekKey();
  return ((meetingSelectionsCache[weekKey] || {})[memberName] || {}).selectedActionIds || [];
}

function saveSelectedActionIds(memberName, ids) {
  const weekKey = getMeetingWeekKey();
  if (!meetingSelectionsCache[weekKey]) meetingSelectionsCache[weekKey] = {};
  if (!meetingSelectionsCache[weekKey][memberName]) meetingSelectionsCache[weekKey][memberName] = {};
  meetingSelectionsCache[weekKey][memberName].selectedActionIds = ids;
  try { localStorage.setItem('meeting-selections-v1-' + weekKey, JSON.stringify(meetingSelectionsCache[weekKey])); } catch(e) {}
  _pushMemberSelectionsToServer(memberName).catch(function() {});
}

function getSelectedStrategyKeys(memberName) {
  const weekKey = getMeetingWeekKey();
  return ((meetingSelectionsCache[weekKey] || {})[memberName] || {}).selectedStrategyKeys || [];
}

function saveSelectedStrategyKeys(memberName, keys) {
  const weekKey = getMeetingWeekKey();
  if (!meetingSelectionsCache[weekKey]) meetingSelectionsCache[weekKey] = {};
  if (!meetingSelectionsCache[weekKey][memberName]) meetingSelectionsCache[weekKey][memberName] = {};
  meetingSelectionsCache[weekKey][memberName].selectedStrategyKeys = keys;
  try { localStorage.setItem('meeting-selections-v1-' + weekKey, JSON.stringify(meetingSelectionsCache[weekKey])); } catch(e) {}
  _pushMemberSelectionsToServer(memberName).catch(function() {});
}

function removeMeetingOgsmItem(memberName, actionId) {
  const ids = getSelectedActionIds(memberName).filter(function(id) { return id !== actionId; });
  saveSelectedActionIds(memberName, ids);
  renderMeetingRows();
}

function removeMeetingStrategyItem(memberName, stratKey) {
  const keys = getSelectedStrategyKeys(memberName).filter(function(k) { return k !== stratKey; });
  saveSelectedStrategyKeys(memberName, keys);
  renderMeetingRows();
}

function renderMeetingStatusFilters() {
  const el = document.getElementById('meeting-status-filters');
  if (!el) return;
  const statuses = ['未開始', '進行中', '卡關', '完成'];
  const members = getMeetingOrderedMembers();
  const counts = { '未開始': 0, '進行中': 0, '卡關': 0, '完成': 0 };
  members.forEach(function(name) {
    const data = name === currentStaff ? state : (staffDataCache[name] || {});
    (data.actions || []).forEach(function(a) {
      if (a.action_name && counts[a.status] !== undefined) counts[a.status]++;
    });
  });
  el.innerHTML = statuses.map(function(s) {
    return '<div class="meeting-status-card meeting-status-card-' + s + '">' +
      '<div class="meeting-status-card-label">' + s + '</div>' +
      '<div class="meeting-status-card-count">' + counts[s] + '</div>' +
    '</div>';
  }).join('');
}

function selectMeetingStatusFilter(status) {}

function toggleMeetingMember(name) {
  meetingCollapsedMembers[name] = !meetingCollapsedMembers[name];
  renderMeetingRows();
}

function renderMeetingRows() {
  const container = document.getElementById('meeting-rows');
  if (!container) return;
  const members = getMeetingOrderedMembers();
  let html = '';
  members.forEach(function(name) {
    const data = name === currentStaff ? state : (staffDataCache[name] || {});
    const allActions = (data.actions || []).filter(function(a) { return !!a.action_name; });
    const selectedIds = getSelectedActionIds(name);
    const selectedActions = selectedIds.map(function(id) {
      return allActions.find(function(a) { return String(a.id) === String(id); });
    }).filter(Boolean);

    const color = avatarColor(name);

    const selectedStrategyKeys = getSelectedStrategyKeys(name);
    const totalSelected = selectedIds.length + selectedStrategyKeys.length;
    const safeName = name.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const pickLabel = totalSelected > 0 ? '編輯項目' : '選取項目';

    const selectedStrategyItems = selectedStrategyKeys.map(function(key) {
      const sep = key.indexOf('::');
      const goalId = key.slice(0, sep);
      const stratName = key.slice(sep + 2);
      const goal = (data.goals || []).find(function(g) { return String(g.id) === goalId; });
      return { key: key, stratName: stratName, goalName: goal ? goal.name : '' };
    });

    let bodyHtml;
    if (selectedActions.length === 0 && selectedStrategyItems.length === 0) {
      bodyHtml = '<div class="meeting-member-empty">尚未選取本週項目</div>';
    } else {
      const stratCards = selectedStrategyItems.map(function(s) {
        const safeKey = s.key.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        return '<div class="meeting-ogsm-card">' +
          '<button class="meeting-ogsm-card-delete" onclick="removeMeetingStrategyItem(\'' + safeName + '\',\'' + safeKey + '\')" title="移除">✕</button>' +
          '<div style="font-size:10px;font-weight:700;color:#0a7a60;margin-bottom:2px">S</div>' +
          '<div class="meeting-ogsm-card-name">' + escHtml(s.stratName) + '</div>' +
        '</div>';
      }).join('');
      const actionCards = selectedActions.map(function(a) {
        const st = a.status || '未開始';
        const safeId = (a.id + '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        return '<div class="meeting-ogsm-card">' +
          '<button class="meeting-ogsm-card-delete" onclick="removeMeetingOgsmItem(\'' + safeName + '\',\'' + safeId + '\')" title="移除">✕</button>' +
          '<div class="meeting-ogsm-card-name">' + escHtml(a.action_name) + '</div>' +
          '<div class="meeting-ogsm-card-footer">' +
            '<span class="mstatus-badge badge-' + escHtml(st) + '">' + escHtml(st) + '</span>' +
            (a.assignee ? '<span class="meeting-ogsm-card-assignee">' + escHtml(a.assignee) + '</span>' : '') +
          '</div>' +
          (a.due_date ? '<span class="meeting-ogsm-card-due">' + escHtml(fmtDate(a.due_date)) + '</span>' : '') +
        '</div>';
      }).join('');
      bodyHtml = '<div class="meeting-ogsm-cards">' + stratCards + actionCards + '</div>';
    }

    const weekKeyNow = getMeetingWeekKey();
    const _mwRangeStr = weekKeyNow + '~' + isoDate(getWeekEnd(getWeekStart(meetingWeekOffset)));
    const noteContent = weekNoteCache[name + '-' + _mwRangeStr] || '';
    const noteId = 'mmn-' + name.replace(/[^a-zA-Z0-9]/g, '_');
    const noteAreaHtml =
      '<div class="meeting-member-note-wrap">' +
        '<div class="meeting-member-note-top">' +
          '<span class="meeting-member-note-label">成果/問題</span>' +
          '<div class="meeting-member-note-toolbar">' +
            '<button class="meeting-member-note-btn" onmousedown="event.preventDefault();meetingMemberNoteCmd(\'' + safeName + '\',\'bold\')" title="粗體"><b>B</b></button>' +
            '<button class="meeting-member-note-btn" onmousedown="event.preventDefault();meetingMemberNoteCmd(\'' + safeName + '\',\'insertUnorderedList\')" title="列點"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4" cy="6" r="1.5" fill="currentColor" stroke="none"/><circle cx="4" cy="12" r="1.5" fill="currentColor" stroke="none"/><circle cx="4" cy="18" r="1.5" fill="currentColor" stroke="none"/></svg></button>' +
            '<button class="meeting-member-note-btn" onmousedown="event.preventDefault();meetingMemberNoteCmd(\'' + safeName + '\',\'link\')" title="超連結"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></button>' +
          '</div>' +
        '</div>' +
        '<div class="meeting-member-note-editor" id="' + noteId + '" contenteditable="true" data-member="' + escHtml(name) + '" data-placeholder="記錄本週成果或發現的問題..." oninput="scheduleMeetingMemberNoteSave(\'' + safeName + '\')">' + noteContent + '</div>' +
      '</div>';

    html += '<div class="meeting-member-section">' +
      '<div class="meeting-member-header">' +
        '<div class="mrow-avatar" style="background:' + color + '">' + escHtml(name[0] || '') + '</div>' +
        '<div class="meeting-member-name">' + escHtml(name) + '</div>' +
        '<span class="meeting-member-count">' + totalSelected + ' 項</span>' +
        '<button class="meeting-pick-btn" onclick="openAiMeetingModal(\'' + safeName + '\')">' + pickLabel + '</button>' +
      '</div>' +
      '<div class="meeting-member-body">' + bodyHtml + noteAreaHtml + '</div>' +
    '</div>';
  });
  container.innerHTML = html || '<div class="meeting-ogsm-hint">無成員資料</div>';
}

function meetingMemberNoteCmd(name, cmd) {
  const noteId = 'mmn-' + name.replace(/[^a-zA-Z0-9]/g, '_');
  const editor = document.getElementById(noteId);
  if (!editor) return;
  editor.focus();
  if (cmd === 'link') {
    showLinkPopover(editor, function(url, displayText, hasSelection) {
      if (hasSelection) {
        document.execCommand('createLink', false, url);
      } else {
        const text = displayText || url;
        document.execCommand('insertHTML', false, '<a href="' + url + '" target="_blank" rel="noopener">' + text + '</a>');
      }
      scheduleMeetingMemberNoteSave(name);
    });
  } else {
    document.execCommand(cmd, false, null);
    scheduleMeetingMemberNoteSave(name);
  }
}

function scheduleMeetingMemberNoteSave(name) {
  const noteId = 'mmn-' + name.replace(/[^a-zA-Z0-9]/g, '_');
  const editor = document.getElementById(noteId);
  if (!editor) return;
  const mws = getWeekStart(meetingWeekOffset);
  const weekRangeStr = isoDate(mws) + '~' + isoDate(getWeekEnd(mws));
  const cacheKey = name + '-' + weekRangeStr;
  weekNoteCache[cacheKey] = editor.innerHTML;
  const statsWs = getWeekStart(statsWeekOffset);
  const statsRangeStr = isoDate(statsWs) + '~' + isoDate(getWeekEnd(statsWs));
  if (name === currentStaff && weekRangeStr === statsRangeStr) {
    const statsEditor = document.getElementById('stats-note-editor');
    if (statsEditor) statsEditor.innerHTML = editor.innerHTML;
  }
  clearTimeout(meetingMemberNoteTimers[name]);
  meetingMemberNoteTimers[name] = setTimeout(async function() {
    try {
      await postData({ type: 'save_week_note', staff: name, weekStart: weekRangeStr, content: weekNoteCache[cacheKey] || '' });
    } finally {
      delete meetingMemberNoteTimers[name];
    }
  }, 2000);
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

function openAiMeetingModal(memberName) {
  meetingPickerMember = memberName;
  pickerActiveGoalId = null;
  pickerActiveStrategyKey = null;
  const modal = document.getElementById('meeting-ogsm-picker');
  const titleEl = document.getElementById('meeting-picker-title');
  if (!modal) return;
  titleEl.textContent = memberName + ' — 選取本週項目';
  modal.style.display = 'flex';

  const existingActionIds = getSelectedActionIds(memberName);
  const existingStratKeys = getSelectedStrategyKeys(memberName);

  if (existingActionIds.length > 0 || existingStratKeys.length > 0) {
    meetingPickerChecked = {
      actions: new Set(existingActionIds.map(String)),
      strategies: new Set(existingStratKeys)
    };
  } else {
    const weekStart = isoDate(getWeekStart(meetingWeekOffset));
    const weekEnd = isoDate(getWeekEnd(getWeekStart(meetingWeekOffset)));
    const data = memberName === currentStaff ? state : (staffDataCache[memberName] || {});
    const allActions = (data.actions || []).filter(function(a) { return !!a.action_name; });
    const suggested = new Set(allActions.filter(function(a) {
      if (a.status === '完成') return false;
      const inWeek = a.due_date && a.due_date >= weekStart && a.due_date <= weekEnd;
      return inWeek || a.status === '進行中';
    }).map(function(a) { return String(a.id); }));
    meetingPickerChecked = { actions: suggested, strategies: new Set() };
  }

  renderPickerModal(memberName);
}

function renderPickerModal(memberName) {
  const bodyEl = document.getElementById('meeting-picker-body');
  if (!bodyEl) return;
  const data = memberName === currentStaff ? state : (staffDataCache[memberName] || {});
  const allGoals = data.goals || [];
  const allStrategies = data.strategies || [];
  const allActions = (data.actions || []).filter(function(a) { return !!a.action_name; });

  if (!allActions.length && !allStrategies.length) {
    bodyEl.innerHTML = '<div class="ai-items-empty">此成員尚無 OGSM 資料</div>';
    return;
  }

  const safeMember = escHtml(memberName).replace(/'/g, "\\'");

  // G column
  let gHtml = '';
  allGoals.forEach(function(goal) {
    const hasItems = allStrategies.some(function(s) { return s.goal_id === goal.id; }) ||
                     allActions.some(function(a) { return a.goal_id === goal.id; });
    if (!hasItems) return;
    const isActive = String(goal.id) === String(pickerActiveGoalId);
    gHtml += '<div class="picker-goal-item' + (isActive ? ' active' : '') +
      '" onclick="setPickerGoal(\'' + goal.id + '\',\'' + safeMember + '\')">' +
      escHtml(goal.name) + '</div>';
  });

  // S column – only show when a G is selected
  let sHtml = '';
  if (!pickerActiveGoalId) {
    sHtml = '<div class="picker-col-empty">← 請先選擇支線目標</div>';
  } else {
    const visibleStrategies = allStrategies.filter(function(s) { return String(s.goal_id) === String(pickerActiveGoalId); });
    if (!visibleStrategies.length) {
      sHtml = '<div class="picker-col-empty">此目標無策略</div>';
    } else {
      visibleStrategies.forEach(function(strat) {
        const stratKey = strat.goal_id + '::' + strat.name;
        const isActive = stratKey === pickerActiveStrategyKey;
        const safeKey = stratKey.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        sHtml += '<div class="picker-goal-item' + (isActive ? ' active' : '') +
          '" onclick="setPickerStrategy(\'' + safeKey + '\',\'' + safeMember + '\')">' +
          escHtml(strat.name) + '</div>';
      });
    }
  }

  // M column – only show when a S is selected
  let mHtml = '';
  if (!pickerActiveStrategyKey) {
    mHtml = '<div class="picker-col-empty">← 請先選擇策略</div>';
  } else {
    const [activeGoalIdStr, activeStratName] = pickerActiveStrategyKey.split('::');
    const visibleActions = allActions.filter(function(a) {
      return String(a.goal_id) === String(activeGoalIdStr) && a.strategy_name === activeStratName;
    });
    if (!visibleActions.length) {
      mHtml = '<div class="picker-col-empty">此策略無行動項目</div>';
    } else {
      visibleActions.forEach(function(a) {
        const isChecked = meetingPickerChecked.actions.has(String(a.id));
        const safeId = (a.id + '').replace(/'/g, "\\'");
        const st = a.status || '未開始';
        mHtml += '<div class="picker-board-item' + (isChecked ? ' active' : '') + '">' +
          '<input type="checkbox" class="picker-checkbox"' + (isChecked ? ' checked' : '') +
          ' onchange="togglePickerItem(\'action\',\'' + safeId + '\');this.closest(\'.picker-board-item\').classList.toggle(\'active\')">' +
          '<div class="picker-board-item-content">' +
            '<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1">' + escHtml(a.action_name) + '</span>' +
            '<span class="mstatus-badge badge-' + escHtml(st) + ' picker-item-badge">' + escHtml(st) + '</span>' +
          '</div>' +
        '</div>';
      });
    }
  }

  bodyEl.innerHTML = '<div class="picker-board">' +
    '<div class="picker-board-col">' +
      '<div class="picker-col-header"><span class="col-tag col-tag-g">G</span>支線目標</div>' +
      '<div class="picker-col-body">' + (gHtml || '<div class="picker-col-empty">無目標</div>') + '</div>' +
    '</div>' +
    '<div class="picker-board-col">' +
      '<div class="picker-col-header"><span class="col-tag col-tag-s">S</span>策略</div>' +
      '<div class="picker-col-body">' + sHtml + '</div>' +
    '</div>' +
    '<div class="picker-board-col">' +
      '<div class="picker-col-header"><span class="col-tag col-tag-m">M</span>Action-行動計劃</div>' +
      '<div class="picker-col-body">' + mHtml + '</div>' +
    '</div>' +
  '</div>';
}

function setPickerGoal(goalId, memberName) {
  if (pickerActiveGoalId === goalId) {
    pickerActiveGoalId = null;
    pickerActiveStrategyKey = null;
  } else {
    pickerActiveGoalId = goalId;
    pickerActiveStrategyKey = null;
  }
  renderPickerModal(memberName);
}

function setPickerStrategy(stratKey, memberName) {
  pickerActiveStrategyKey = pickerActiveStrategyKey === stratKey ? null : stratKey;
  renderPickerModal(memberName);
}

function togglePickerItem(type, id) {
  if (type === 'action') {
    if (meetingPickerChecked.actions.has(String(id))) meetingPickerChecked.actions.delete(String(id));
    else meetingPickerChecked.actions.add(String(id));
  } else {
    if (meetingPickerChecked.strategies.has(id)) meetingPickerChecked.strategies.delete(id);
    else meetingPickerChecked.strategies.add(id);
  }
}

function closeAiMeetingModal() {
  const modal = document.getElementById('meeting-ogsm-picker');
  if (modal) modal.style.display = 'none';
  meetingPickerMember = null;
  aiMeetingTempItems = [];
  meetingPickerChecked = { actions: new Set(), strategies: new Set() };
  pickerActiveGoalId = null;
  pickerActiveStrategyKey = null;
}

async function confirmAiMeetingItems() {
  if (!meetingPickerMember) return;
  const memberToSync = meetingPickerMember;
  const weekKey = getMeetingWeekKey();
  const selectedActionIds = Array.from(meetingPickerChecked.actions);
  const selectedStrategyKeys = Array.from(meetingPickerChecked.strategies);
  if (!meetingSelectionsCache[weekKey]) meetingSelectionsCache[weekKey] = {};
  meetingSelectionsCache[weekKey][memberToSync] = { selectedActionIds: selectedActionIds, selectedStrategyKeys: selectedStrategyKeys };
  try { localStorage.setItem('meeting-selections-v1-' + weekKey, JSON.stringify(meetingSelectionsCache[weekKey])); } catch(e) {}
  closeAiMeetingModal();
  renderMeetingRows();
  try {
    await _pushMemberSelectionsToServer(memberToSync);
  } catch(e) {
    await new Promise(function(r) { setTimeout(r, 2000); });
    try {
      await _pushMemberSelectionsToServer(memberToSync);
    } catch(e2) {
      showToast('❌ 同步失敗，請重試', true);
    }
  }
}

function switchMeetingTab(tab) {
  ['report', 'announce'].forEach(function(t) {
    const panel = document.getElementById('meeting-tab-' + t);
    const btn = document.getElementById('mtab-' + t);
    if (panel) panel.style.display = (t === tab) ? '' : 'none';
    if (btn) btn.classList.toggle('active', t === tab);
  });
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
    const cached = (meetingMemberNotesCache[weekKey] || {})._announce;
    const content = cached !== undefined ? cached : (localStorage.getItem('meeting-announce-' + weekKey) || '');
    editor.innerHTML = content;
    if (!editor._linkClickInited) {
      editor._linkClickInited = true;
      editor.addEventListener('click', function(e) {
        const a = e.target.closest('a');
        if (a) { e.preventDefault(); window.open(a.href, '_blank'); }
      });
    }
  }
  renderMeetingAnnounceHistory();
}

function meetingAnnounceCmd(cmd) {
  const editor = document.getElementById('meeting-announce-editor');
  if (!editor) return;
  editor.focus();
  if (cmd === 'link') {
    showLinkPopover(editor, function(url, displayText, hasSelection) {
      if (hasSelection) {
        document.execCommand('createLink', false, url);
        editor.querySelectorAll('a[href="' + url + '"]').forEach(function(a) {
          a.setAttribute('target', '_blank');
          a.setAttribute('rel', 'noopener');
        });
      } else {
        const text = displayText || url;
        document.execCommand('insertHTML', false, '<a href="' + url + '" target="_blank" rel="noopener">' + text + '</a>');
      }
    });
  } else {
    document.execCommand(cmd, false, null);
  }
}

function meetingAnnounceSave() {
  const editor = document.getElementById('meeting-announce-editor');
  if (!editor) return;
  const weekKey = getMeetingWeekKey();
  const content = editor.innerHTML;
  localStorage.setItem('meeting-announce-' + weekKey, content);
  if (!meetingMemberNotesCache[weekKey]) meetingMemberNotesCache[weekKey] = {};
  meetingMemberNotesCache[weekKey]._announce = content;
  postData({ type: 'save_meeting_note', noteType: 'announce', weekKey: weekKey, member: '', content: content }).catch(function() {});
  showToast('✅ 佈達事項已儲存');
  renderMeetingAnnounceHistory();
}

async function renderMeetingAnnounceHistory() {
  const listEl = document.getElementById('meeting-announce-history-list');
  if (!listEl) return;

  const currentWeekKey = getMeetingWeekKey();
  const currentKey = 'meeting-announce-' + currentWeekKey;

  function renderList(items) {
    if (!items.length) {
      listEl.innerHTML = '<div class="announce-history-empty">尚無歷史紀錄</div>';
      return;
    }
    listEl.innerHTML = items.map(function(item) {
      return '<details class="announce-history-item">' +
        '<summary class="announce-history-summary">' + escHtml(item.weekKey) + ' 週</summary>' +
        '<div class="announce-history-content">' + item.content + '</div>' +
        '</details>';
    }).join('');
  }

  // Show localStorage history first for instant feedback
  const localKeys = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith('meeting-announce-') && k !== currentKey) localKeys.push(k);
  }
  localKeys.sort().reverse();
  renderList(localKeys.map(function(k) {
    return { weekKey: k.replace('meeting-announce-', ''), content: localStorage.getItem(k) || '' };
  }).filter(function(i) { return !!i.content; }));

  // Fetch from backend and update
  try {
    const h = { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/meeting_notes?note_type=eq.announce&week_key=neq.${encodeURIComponent(currentWeekKey)}&select=week_key,content&order=week_key.desc`, { headers: h });
    const data = await res.json();
    if (data.length > 0) {
      const history = data.filter(function(i) { return !!i.content; }).map(function(i) { return { weekKey: i.week_key, content: i.content }; });
      history.forEach(function(item) {
        if (item.content) localStorage.setItem('meeting-announce-' + item.weekKey, item.content);
      });
      renderList(history);
    }
  } catch(e) {}
}
