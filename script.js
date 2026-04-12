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
let currentStaff = localStorage.getItem('ogsm-current-staff') || 'Riku';
let staffList    = [];

// ── Fetch / Post ──
async function fetchData() {
  const res = await fetch(GAS_URL + '?api=1&staff=' + encodeURIComponent(currentStaff), { method: 'GET' });
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
      const res = await postData({ type:'rename_objective', obj_id:obj.id, new_title:newTitle });
      if (res.success) { obj.title = newTitle; state.objectives[0] = obj; showToast('✅ 目的已更新'); }
      else { showToast('❌ '+(res.message||'更新失敗'), true); nameEl.textContent = obj.title; }
    } catch(e) { showToast('❌ 網路錯誤', true); nameEl.textContent = obj.title; }
  };
  nameEl.onkeydown = function(e) {
    if (e.key==='Enter') { e.preventDefault(); nameEl.blur(); }
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
      item.innerHTML = `
        <div class="goal-item-top-row">
          <div class="goal-item-num">目標 ${idx+1}</div>
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
        if (e.key==='Enter') { e.preventDefault(); nameEl.blur(); }
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
      item.innerHTML = `
        <div class="strategy-item-top-row">
          <div style="display:flex;align-items:center;gap:6px">
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
        const statuses = ['進行中','需要協助','受阻','成功'];
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
        if (e.key==='Enter') { e.preventDefault(); sNameEl.blur(); }
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
      item.innerHTML = `
        <div class="action-item-top">
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
      aNameEl.addEventListener('keydown', function(e) {
        if (e.key==='Enter') { e.preventDefault(); aNameEl.blur(); }
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

function makeColumn(tag, title, tagClass, count) {
  const col = document.createElement('div');
  col.className = 'col-panel';
  col.innerHTML = `
    <div class="col-header">
      <span class="col-tag ${tagClass}">${tag}</span>
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
    if (res.success) { showToast('✅ 更新成功'); closeEditModal(); await loadAndRender(); }
    else showToast('❌ '+(res.message||'更新失敗'), true);
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
  return getStrategyData(goalId, stratName).status || '進行中';
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
  staffList.forEach(name => {
    const chip = document.createElement('button');
    chip.className = 'staff-chip' + (name === currentStaff ? ' active' : '');
    chip.innerHTML = `<span class="staff-chip-label">${escHtml(name)}</span><span class="staff-chip-del" title="刪除職員">✕</span>`;
    chip.addEventListener('click', function(e) {
      if (e.target.classList.contains('staff-chip-del')) {
        e.stopPropagation();
        openDeleteStaffConfirm(name);
      } else {
        switchStaff(name);
      }
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
  await loadAndRender();
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

// ── Main ──
async function loadAndRender() {
  try {
    const data = await fetchData();
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
}

init();
