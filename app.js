/* ===================================================
   劳动案件管理系统 — app.js
   数据存储：localStorage (lc_cases / lc_parties)
=================================================== */

// ===== Storage =====
const Storage = {
  getCases() {
    try { return JSON.parse(localStorage.getItem('lc_cases') || '[]'); }
    catch { return []; }
  },
  saveCases(cases) {
    localStorage.setItem('lc_cases', JSON.stringify(cases));
  },
  getParties() {
    try { return JSON.parse(localStorage.getItem('lc_parties') || '[]'); }
    catch { return []; }
  },
  saveParties(parties) {
    localStorage.setItem('lc_parties', JSON.stringify(parties));
  },
  getHandlers() {
    try { return JSON.parse(localStorage.getItem('lc_handlers') || '[]'); }
    catch { return []; }
  },
  saveHandlers(handlers) {
    localStorage.setItem('lc_handlers', JSON.stringify(handlers));
  }
};

// ===== Utils =====
function uuid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

function urgencyLevel(courtDate) {
  if (!courtDate) return 'gray';
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const court = new Date(courtDate); court.setHours(0, 0, 0, 0);
  const diff = Math.ceil((court - today) / (1000 * 60 * 60 * 24));
  if (diff <= 15) return 'red';
  if (diff <= 60) return 'yellow';
  return 'green';
}

// 举证期限截止日 = 开庭日期 - evidenceDays 天
function calcEvidenceDeadline(courtDate, days) {
  if (!courtDate || !days || isNaN(days)) return null;
  const d = new Date(courtDate);
  d.setDate(d.getDate() - parseInt(days));
  return d.toISOString().slice(0, 10); // YYYY-MM-DD
}

function formatDate(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr);
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
}

function statusBadge(status) {
  const map = { '进行中': 'badge-active', '已结案': 'badge-closed' };
  return `<span class="badge ${map[status] || 'badge-active'}">${status || '进行中'}</span>`;
}

function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function urgencyLabel(level) {
  return { red: '紧急（≤15天）', yellow: '注意（16-60天）', green: '正常（>60天）', gray: '无开庭日期' }[level] || '';
}

// ===== State =====
let currentFilter = 'all';
let currentSearch = '';
let pendingAttachments = [];

// ===== Render Cards =====
function getMyPartiesList(c) {
  const parties = Storage.getParties();
  if (c.myParties && c.myParties.length > 0) {
    return c.myParties.map(mp => {
      if (mp.partyId) {
        const p = parties.find(p => p.id === mp.partyId);
        return p ? p.companyName : mp.custom || '';
      }
      return mp.custom || '';
    }).filter(Boolean);
  }
  // backward compat
  if (c.myPartyId) {
    const p = parties.find(p => p.id === c.myPartyId);
    return [p ? p.companyName : (c.myPartyCustom || '')].filter(Boolean);
  }
  return c.myPartyCustom ? [c.myPartyCustom] : [];
}

function getDisplayPartyName(c) {
  const list = getMyPartiesList(c);
  return list.join('；') || '—';
}

function getPartyOptionsHtml() {
  const parties = Storage.getParties();
  return '<option value="">-- 从主体库选择 --</option>' +
    parties.map(p => `<option value="${p.id}">${escHtml(p.companyName)}</option>`).join('');
}

function addPartyRow(partyId, custom) {
  const container = document.getElementById('myPartiesContainer');
  const row = document.createElement('div');
  row.className = 'party-row';
  row.innerHTML = `
    <select class="my-party-select">${getPartyOptionsHtml()}</select>
    <span class="party-or">或</span>
    <input type="text" class="my-party-custom" placeholder="手动输入我方主体" />
    <button type="button" class="btn-icon remove-party-row" title="删除">✕</button>`;
  if (partyId) row.querySelector('.my-party-select').value = partyId;
  if (custom) row.querySelector('.my-party-custom').value = custom;
  row.querySelector('.remove-party-row').addEventListener('click', () => {
    row.remove();
    refreshPartyRemoveBtns();
  });
  container.appendChild(row);
  refreshPartyRemoveBtns();
}

function refreshPartyRemoveBtns() {
  const rows = document.querySelectorAll('#myPartiesContainer .party-row');
  rows.forEach(row => {
    row.querySelector('.remove-party-row').style.display = rows.length === 1 ? 'none' : '';
  });
}

function renderCards() {
  const cases = Storage.getCases();
  const grid = document.getElementById('cardsGrid');
  const empty = document.getElementById('emptyState');

  let filtered = cases;

  if (currentFilter !== 'all') {
    filtered = filtered.filter(c => c.status === currentFilter);
  }

  if (currentSearch.trim()) {
    const q = currentSearch.trim().toLowerCase();
    filtered = filtered.filter(c => {
      const partyName = getDisplayPartyName(c).toLowerCase();
      return (
        (c.name || '').toLowerCase().includes(q) ||
        (c.causeOfAction || '').toLowerCase().includes(q) ||
        (c.oppositeParty || '').toLowerCase().includes(q) ||
        partyName.includes(q)
      );
    });
  }

  // Sort: court date asc, no date goes to end
  filtered.sort((a, b) => {
    if (a.courtDate && b.courtDate) return new Date(a.courtDate) - new Date(b.courtDate);
    if (a.courtDate) return -1;
    if (b.courtDate) return 1;
    return 0;
  });

  if (filtered.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  grid.innerHTML = filtered.map(c => {
    const level = urgencyLevel(c.courtDate);
    const partyName = getDisplayPartyName(c);
    const evidenceDeadlineDate = calcEvidenceDeadline(c.courtDate, c.evidenceDays);
    return `
      <div class="case-card" data-id="${c.id}">
        <div class="card-header">
          <span class="urgency-dot urgency-${level}" title="${urgencyLabel(level)}"></span>
          <span class="card-title">${escHtml(c.name)}</span>
        </div>
        <div class="card-meta">
          <div class="card-meta-row"><span class="meta-label">案由</span><span class="meta-value">${escHtml(c.causeOfAction) || ''}</span></div>
          <div class="card-meta-row"><span class="meta-label">我方</span><span class="meta-value">${escHtml(partyName) || ''}</span></div>
          <div class="card-meta-row"><span class="meta-label">开庭地</span><span class="meta-value">${escHtml(c.courtLocation) || ''}</span></div>
          <div class="card-meta-row"><span class="meta-label">举证截止</span><span class="meta-value">${evidenceDeadlineDate ? formatDate(evidenceDeadlineDate) : ''}</span></div>
          <div class="card-meta-row"><span class="meta-label">开庭日</span><span class="meta-value">${formatDate(c.courtDate)}</span></div>
          <div class="card-meta-row"><span class="meta-label">办案人</span><span class="meta-value">${escHtml(c.handler) || ''}</span></div>
        </div>
        <div class="card-footer">
          <span class="badge badge-stage">${escHtml(c.stage || '仲裁')}</span>
          ${c.caseType ? `<span class="badge badge-type">${escHtml(c.caseType)}</span>` : ''}
          ${statusBadge(c.status)}
        </div>
      </div>
    `;
  }).join('');

  grid.querySelectorAll('.case-card').forEach(card => {
    card.addEventListener('click', () => openCaseModal(card.dataset.id));
  });
}

// ===== Case Modal =====
function openCaseModal(id) {
  const modal = document.getElementById('caseModal');
  const form = document.getElementById('caseForm');

  form.reset();
  pendingAttachments = [];
  renderAttachmentList();
  clearErrors();
  document.getElementById('evidenceHint').textContent = '';

  // Reset party rows
  document.getElementById('myPartiesContainer').innerHTML = '';

  // Show delete button only when editing
  document.getElementById('deleteCaseBtn').style.display = id ? '' : 'none';

  if (id) {
    const cases = Storage.getCases();
    const c = cases.find(x => x.id === id);
    if (!c) return;
    document.getElementById('caseModalTitle').textContent = '编辑案件';
    document.getElementById('caseId').value = c.id;
    document.getElementById('caseName').value = c.name || '';
    document.getElementById('caseNumber').value = c.caseNumber || '';
    document.getElementById('causeOfAction').value = c.causeOfAction || '';
    // Populate party rows (new format or backward compat)
    if (c.myParties && c.myParties.length > 0) {
      c.myParties.forEach(mp => addPartyRow(mp.partyId || '', mp.custom || ''));
    } else {
      addPartyRow(c.myPartyId || '', c.myPartyCustom || '');
    }
    document.getElementById('oppositeParty').value = c.oppositeParty || '';
    document.getElementById('courtLocation').value = c.courtLocation || '';
    document.getElementById('courtDate').value = c.courtDate || '';
    document.getElementById('evidenceDays').value = c.evidenceDays || '';
    document.getElementById('claims').value = c.claims || '';
    document.getElementById('handler').value = c.handler || '';
    document.getElementById('caseStatus').value = c.status || '进行中';
    document.getElementById('caseType').value = c.caseType || '';
    const stageEl = form.querySelector(`input[name="stage"][value="${c.stage}"]`);
    if (stageEl) stageEl.checked = true;
    pendingAttachments = (c.attachments || []).map(name => ({ name }));
    renderAttachmentList();
    updateEvidenceHint();
  } else {
    document.getElementById('caseModalTitle').textContent = '新建案件';
    document.getElementById('caseId').value = '';
    addPartyRow('', '');
  }

  modal.style.display = 'flex';
}

function closeCaseModal() {
  document.getElementById('caseModal').style.display = 'none';
}

function updateEvidenceHint() {
  const courtDate = document.getElementById('courtDate').value;
  const days = document.getElementById('evidenceDays').value;
  const hint = document.getElementById('evidenceHint');
  const deadline = calcEvidenceDeadline(courtDate, days);
  hint.textContent = deadline ? `截止日：${formatDate(deadline)}` : '';
}

function saveCase() {
  clearErrors();
  const name = document.getElementById('caseName').value.trim();
  if (!name) {
    document.getElementById('caseName').classList.add('error');
    document.getElementById('caseNameErr').textContent = '案件名称不能为空';
    document.getElementById('caseName').focus();
    return;
  }

  const id = document.getElementById('caseId').value;
  const stage = document.querySelector('input[name="stage"]:checked')?.value || '仲裁';
  const evidenceDaysVal = document.getElementById('evidenceDays').value;

  const myParties = [];
  document.querySelectorAll('#myPartiesContainer .party-row').forEach(row => {
    const partyId = row.querySelector('.my-party-select').value;
    const custom = row.querySelector('.my-party-custom').value.trim();
    if (partyId || custom) myParties.push({ partyId: partyId || null, custom: partyId ? '' : custom });
  });

  const data = {
    name,
    caseNumber: document.getElementById('caseNumber').value.trim(),
    causeOfAction: document.getElementById('causeOfAction').value.trim(),
    myParties,
    oppositeParty: document.getElementById('oppositeParty').value.trim(),
    courtLocation: document.getElementById('courtLocation').value.trim(),
    courtDate: document.getElementById('courtDate').value,
    evidenceDays: evidenceDaysVal ? parseInt(evidenceDaysVal) : null,
    claims: document.getElementById('claims').value.trim(),
    stage,
    handler: document.getElementById('handler').value.trim(),
    status: document.getElementById('caseStatus').value,
    caseType: document.getElementById('caseType').value,
    attachments: pendingAttachments.map(a => a.name),
    updatedAt: new Date().toISOString()
  };

  const cases = Storage.getCases();
  if (id) {
    const idx = cases.findIndex(c => c.id === id);
    if (idx >= 0) cases[idx] = { ...cases[idx], ...data };
  } else {
    data.id = uuid();
    data.createdAt = new Date().toISOString();
    cases.unshift(data);
  }

  Storage.saveCases(cases);
  closeCaseModal();
  renderCards();
}

function clearErrors() {
  document.getElementById('caseName').classList.remove('error');
  document.getElementById('caseNameErr').textContent = '';
}

// ===== Attachments =====
function renderAttachmentList() {
  const list = document.getElementById('attachmentList');
  list.innerHTML = pendingAttachments.map((a, i) => `
    <li class="attachment-item">
      <span>📄 ${escHtml(a.name)}</span>
      <button class="remove-attach" data-idx="${i}" title="移除">✕</button>
    </li>
  `).join('');
  list.querySelectorAll('.remove-attach').forEach(btn => {
    btn.addEventListener('click', () => {
      pendingAttachments.splice(parseInt(btn.dataset.idx), 1);
      renderAttachmentList();
    });
  });
}

// ===== Party Modal =====
function openPartyModal() {
  showPartyList();
  document.getElementById('partyModal').style.display = 'flex';
}

function closePartyModal() {
  document.getElementById('partyModal').style.display = 'none';
}

function showPartyList() {
  document.getElementById('partyListSection').style.display = 'block';
  document.getElementById('partyFormSection').style.display = 'none';
  renderPartyTable();
}

function renderPartyTable() {
  const parties = Storage.getParties();
  const tbody = document.getElementById('partyTableBody');
  const emptyEl = document.getElementById('emptyParty');
  const table = document.getElementById('partyTable');

  if (parties.length === 0) {
    table.style.display = 'none';
    emptyEl.style.display = 'block';
    return;
  }
  table.style.display = '';
  emptyEl.style.display = 'none';

  tbody.innerHTML = parties.map(p => `
    <tr>
      <td>${escHtml(p.companyName)}</td>
      <td>${escHtml(p.creditCode) || '—'}</td>
      <td>${escHtml(p.legalRep) || '—'}</td>
      <td>
        <div class="actions">
          <button class="btn btn-outline btn-sm" data-edit="${p.id}">编辑</button>
          <button class="btn btn-ghost btn-sm" data-del="${p.id}">删除</button>
        </div>
      </td>
    </tr>
  `).join('');

  tbody.querySelectorAll('[data-edit]').forEach(btn => {
    btn.addEventListener('click', () => openPartyForm(btn.dataset.edit));
  });
  tbody.querySelectorAll('[data-del]').forEach(btn => {
    btn.addEventListener('click', () => deleteParty(btn.dataset.del));
  });
}

function openPartyForm(id) {
  const form = document.getElementById('partyForm');
  form.reset();
  document.getElementById('partyNameErr').textContent = '';
  document.getElementById('partyId').value = '';

  if (id) {
    const parties = Storage.getParties();
    const p = parties.find(x => x.id === id);
    if (!p) return;
    document.getElementById('partyFormTitle').textContent = '编辑主体';
    document.getElementById('partyId').value = p.id;
    document.getElementById('partyCompanyName').value = p.companyName || '';
    document.getElementById('partyCreditCode').value = p.creditCode || '';
    document.getElementById('partyLegalRep').value = p.legalRep || '';
    document.getElementById('partyAddress').value = p.address || '';
  } else {
    document.getElementById('partyFormTitle').textContent = '新增主体';
  }

  document.getElementById('partyListSection').style.display = 'none';
  document.getElementById('partyFormSection').style.display = 'block';
}

function saveParty() {
  const companyName = document.getElementById('partyCompanyName').value.trim();
  if (!companyName) {
    document.getElementById('partyNameErr').textContent = '公司名称不能为空';
    document.getElementById('partyCompanyName').focus();
    return;
  }
  document.getElementById('partyNameErr').textContent = '';

  const id = document.getElementById('partyId').value;
  const data = {
    companyName,
    creditCode: document.getElementById('partyCreditCode').value.trim(),
    legalRep: document.getElementById('partyLegalRep').value.trim(),
    address: document.getElementById('partyAddress').value.trim()
  };

  const parties = Storage.getParties();
  if (id) {
    const idx = parties.findIndex(p => p.id === id);
    if (idx >= 0) parties[idx] = { ...parties[idx], ...data };
  } else {
    data.id = uuid();
    parties.push(data);
  }

  Storage.saveParties(parties);
  showPartyList();
}

function deleteParty(id) {
  const cases = Storage.getCases();
  const inUse = cases.some(c => c.myPartyId === id);
  if (inUse) {
    alert('该主体已被案件引用，无法删除。请先修改相关案件的我方主体。');
    return;
  }
  if (!confirm('确定删除该主体吗？')) return;
  const parties = Storage.getParties().filter(p => p.id !== id);
  Storage.saveParties(parties);
  renderPartyTable();
}

// ===== Handler Modal =====
function openHandlerModal() {
  showHandlerList();
  document.getElementById('handlerModal').style.display = 'flex';
}

function closeHandlerModal() {
  document.getElementById('handlerModal').style.display = 'none';
}

function showHandlerList() {
  document.getElementById('handlerListSection').style.display = 'block';
  document.getElementById('handlerFormSection').style.display = 'none';
  renderHandlerTable();
}

function renderHandlerTable() {
  const handlers = Storage.getHandlers();
  const tbody = document.getElementById('handlerTableBody');
  const emptyEl = document.getElementById('emptyHandler');
  const table = document.getElementById('handlerTable');

  if (handlers.length === 0) {
    table.style.display = 'none';
    emptyEl.style.display = 'block';
    return;
  }
  table.style.display = '';
  emptyEl.style.display = 'none';

  tbody.innerHTML = handlers.map(h => `
    <tr>
      <td>${escHtml(h.name)}</td>
      <td>${escHtml(h.idCard) || '—'}</td>
      <td>${escHtml(h.phone) || '—'}</td>
      <td>
        <div class="actions">
          <button class="btn btn-outline btn-sm" data-edit="${h.id}">编辑</button>
          <button class="btn btn-ghost btn-sm" data-del="${h.id}">删除</button>
        </div>
      </td>
    </tr>
  `).join('');

  tbody.querySelectorAll('[data-edit]').forEach(btn => {
    btn.addEventListener('click', () => openHandlerForm(btn.dataset.edit));
  });
  tbody.querySelectorAll('[data-del]').forEach(btn => {
    btn.addEventListener('click', () => deleteHandler(btn.dataset.del));
  });
}

function openHandlerForm(id) {
  const form = document.getElementById('handlerForm');
  form.reset();
  document.getElementById('handlerNameErr').textContent = '';
  document.getElementById('handlerId').value = '';

  if (id) {
    const h = Storage.getHandlers().find(x => x.id === id);
    if (!h) return;
    document.getElementById('handlerFormTitle').textContent = '编辑办案人';
    document.getElementById('handlerId').value = h.id;
    document.getElementById('handlerName').value = h.name || '';
    document.getElementById('handlerIdCard').value = h.idCard || '';
    document.getElementById('handlerPhone').value = h.phone || '';
  } else {
    document.getElementById('handlerFormTitle').textContent = '新增办案人';
  }

  document.getElementById('handlerListSection').style.display = 'none';
  document.getElementById('handlerFormSection').style.display = 'block';
}

function saveHandler() {
  const name = document.getElementById('handlerName').value.trim();
  if (!name) {
    document.getElementById('handlerNameErr').textContent = '办案人姓名不能为空';
    document.getElementById('handlerName').focus();
    return;
  }
  document.getElementById('handlerNameErr').textContent = '';

  const id = document.getElementById('handlerId').value;
  const data = {
    name,
    idCard: document.getElementById('handlerIdCard').value.trim(),
    phone: document.getElementById('handlerPhone').value.trim()
  };

  const handlers = Storage.getHandlers();
  if (id) {
    const idx = handlers.findIndex(h => h.id === id);
    if (idx >= 0) handlers[idx] = { ...handlers[idx], ...data };
  } else {
    data.id = uuid();
    handlers.push(data);
  }

  Storage.saveHandlers(handlers);
  showHandlerList();
}

function deleteHandler(id) {
  if (!confirm('确定删除该办案人吗？')) return;
  const handlers = Storage.getHandlers().filter(h => h.id !== id);
  Storage.saveHandlers(handlers);
  renderHandlerTable();
}

// ===== Batch Import =====
// 列定义：key, 表头文字, 是否必填, 类型(text/date/number/select), 选项
const BATCH_COLS = [
  { key: 'name',          label: '案件名称',     required: true,  type: 'text' },
  { key: 'causeOfAction', label: '案由',         required: false, type: 'text' },
  { key: 'myPartyCustom', label: '我方主体',     required: false, type: 'text' },
  { key: 'caseType',      label: '案件类型',     required: false, type: 'text',   tooltip: '雇员 / 非雇员' },
  { key: 'oppositeParty', label: '对方主体',     required: false, type: 'text' },
  { key: 'courtLocation', label: '开庭地点',     required: false, type: 'text' },
  { key: 'courtDate',     label: '开庭日期',     required: false, type: 'text',   hint: 'YYYY-MM-DD' },
  { key: 'evidenceDays',  label: '举证期限(天)', required: false, type: 'text',   hint: '数字' },
  { key: 'stage',         label: '案件阶段',     required: false, type: 'text',   tooltip: '仲裁 / 一审 / 二审 / 再审', defaultVal: '仲裁' },
  { key: 'handler',       label: '办案人',       required: false, type: 'text' },
  { key: 'claims',        label: '诉讼请求',     required: false, type: 'text' },
  { key: 'status',        label: '案件状态',     required: false, type: 'text',   tooltip: '进行中 / 已结案', defaultVal: '进行中' },
];

const BATCH_INIT_ROWS = 1;

function openBatchModal() {
  buildBatchTable();
  document.getElementById('batchError').textContent = '';
  document.getElementById('batchModal').style.display = 'flex';
}

function closeBatchModal() {
  document.getElementById('batchModal').style.display = 'none';
}

function buildBatchTable() {
  // Headers
  const headerRow = document.getElementById('batchHeaderRow');
  headerRow.innerHTML = BATCH_COLS.map(col => {
    const hintPart = col.hint ? `<br><small style="font-weight:400;color:#94A3B8">${col.hint}</small>` : '';
    const tooltipPart = col.tooltip
      ? `<span class="col-tooltip-wrap" data-tip="${col.tooltip}">?</span>`
      : '';
    return `<th>${col.label}${col.required ? ' <span class="col-required">*</span>' : ''}${tooltipPart}${hintPart}</th>`;
  }).join('') + '<th></th>';

  // Wire up floating tooltip
  const floatTip = document.getElementById('col-float-tooltip');
  headerRow.querySelectorAll('.col-tooltip-wrap').forEach(el => {
    el.addEventListener('mouseenter', e => {
      floatTip.textContent = el.dataset.tip;
      floatTip.style.display = 'block';
      const r = el.getBoundingClientRect();
      const tipW = floatTip.offsetWidth;
      let left = r.left + r.width / 2 - tipW / 2;
      // clamp to viewport
      left = Math.max(8, Math.min(left, window.innerWidth - tipW - 8));
      floatTip.style.left = left + 'px';
      floatTip.style.top = (r.bottom + 8) + 'px';
    });
    el.addEventListener('mouseleave', () => { floatTip.style.display = 'none'; });
  });

  // Body: reset
  const tbody = document.getElementById('batchTableBody');
  tbody.innerHTML = '';
  for (let i = 0; i < BATCH_INIT_ROWS; i++) addBatchRow();
}

function addBatchRow() {
  const tbody = document.getElementById('batchTableBody');
  const tr = document.createElement('tr');
  tr.innerHTML = BATCH_COLS.map(col => {
    const ph = col.hint || '';
    return `<td><input type="text" class="batch-cell" data-col="${col.key}" placeholder="${ph}" /></td>`;
  }).join('') + `<td><button class="btn-icon remove-batch-row" title="删除行">✕</button></td>`;

  // Paste handler on each cell
  tr.querySelectorAll('.batch-cell').forEach(cell => {
    cell.addEventListener('paste', onBatchCellPaste);
  });
  tr.querySelector('.remove-batch-row').addEventListener('click', () => tr.remove());
  tbody.appendChild(tr);
}

function onBatchCellPaste(e) {
  const text = (e.clipboardData || window.clipboardData).getData('text');
  // Check if it's multi-cell data (contains tabs or newlines)
  if (!text.includes('\t') && !text.includes('\n')) return; // single cell, default paste

  e.preventDefault();

  const rows = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trimEnd().split('\n');
  const tbody = document.getElementById('batchTableBody');
  const allRows = Array.from(tbody.querySelectorAll('tr'));

  // Find starting position
  const currentCell = e.currentTarget;
  const currentRow = currentCell.closest('tr');
  const colIdx = Array.from(currentRow.querySelectorAll('.batch-cell')).indexOf(currentCell);
  let rowIdx = allRows.indexOf(currentRow);

  rows.forEach((rowText, ri) => {
    const cells = rowText.split('\t');
    // Ensure we have enough rows
    while (rowIdx + ri >= tbody.querySelectorAll('tr').length) addBatchRow();
    const targetRow = tbody.querySelectorAll('tr')[rowIdx + ri];
    const targetCells = Array.from(targetRow.querySelectorAll('.batch-cell'));

    cells.forEach((val, ci) => {
      const targetCell = targetCells[colIdx + ci];
      if (!targetCell) return;
      targetCell.value = val.trim();
      targetCell.classList.remove('cell-error');
    });
  });
}

function importBatch() {
  const tbody = document.getElementById('batchTableBody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const errorEl = document.getElementById('batchError');
  errorEl.textContent = '';

  // Clear previous errors
  tbody.querySelectorAll('.cell-error').forEach(el => el.classList.remove('cell-error'));
  rows.forEach(r => r.classList.remove('row-error'));

  const newCases = [];
  let hasError = false;

  rows.forEach((tr, i) => {
    const cells = tr.querySelectorAll('.batch-cell');
    const rowData = {};
    cells.forEach(cell => { rowData[cell.dataset.col] = cell.value.trim(); });

    // Skip completely empty rows
    const allEmpty = BATCH_COLS.every(col => !rowData[col.key]);
    if (allEmpty) return;

    // Validate required
    if (!rowData.name) {
      tr.querySelector('[data-col="name"]').classList.add('cell-error');
      tr.classList.add('row-error');
      hasError = true;
    }

    if (!hasError || rowData.name) {
      const evidenceDaysNum = rowData.evidenceDays ? parseInt(rowData.evidenceDays) : null;
      const batchMyParties = (rowData.myPartyCustom || '')
        .split(/[；;]/).map(s => s.trim()).filter(Boolean)
        .map(s => ({ partyId: null, custom: s }));
      newCases.push({
        id: uuid(),
        name: rowData.name,
        causeOfAction: rowData.causeOfAction || '',
        myParties: batchMyParties,
        oppositeParty: rowData.oppositeParty || '',
        courtLocation: rowData.courtLocation || '',
        courtDate: rowData.courtDate || '',
        evidenceDays: isNaN(evidenceDaysNum) ? null : evidenceDaysNum,
        claims: rowData.claims || '',
        stage: rowData.stage || '仲裁',
        handler: rowData.handler || '',
        status: rowData.status || '进行中',
        attachments: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    }
  });

  if (hasError) {
    errorEl.textContent = '请填写所有标红行的案件名称';
    return;
  }

  if (newCases.length === 0) {
    errorEl.textContent = '请至少填写一行案件数据';
    return;
  }

  const cases = Storage.getCases();
  Storage.saveCases([...newCases, ...cases]);
  closeBatchModal();
  renderCards();
}

// ===== Document Generation =====
const DOC_TEMPLATES = {
  '答辩状': (f) => `
<p style="text-align:center;font-size:18pt;font-weight:bold;letter-spacing:0.5em;margin:0 0 8pt 0;">答辩状</p>
<p style="text-align:center;font-size:12pt;margin:0 0 14pt 0;">案号：${escHtml(f.caseNumber)}</p>
<p style="font-size:12pt;margin:0 0 4pt 0;"><strong>答辩人：</strong>${escHtml(f.companyName)}</p>
<p style="font-size:12pt;margin:0 0 4pt 0;"><strong>住所：</strong>${escHtml(f.address)}</p>
<p style="font-size:12pt;margin:0 0 14pt 0;"><strong>负责人：</strong>${escHtml(f.legalRep)}</p>
<p style="font-size:12pt;text-indent:2em;margin:0 0 12pt 0;">关于被答辩人${escHtml(f.caseName)}一案，答辩人认为：</p>
<p style="font-size:12pt;text-indent:2em;margin:0 0 6pt 0;"><strong>针对诉求一…</strong></p>
<p style="font-size:12pt;text-indent:2em;margin:0 0 20pt 0;"><strong>针对诉求二…</strong></p>
<p style="font-size:12pt;text-indent:2em;margin:0 0 0 0;">此　致</p>
<p style="font-size:12pt;margin:0 0 20pt 0;">${escHtml(f.courtLocation)}</p>
<p style="font-size:12pt;text-align:right;margin:0 0 4pt 0;">答辩人：${escHtml(f.companyName)}</p>
<p style="font-size:12pt;text-align:right;margin:0;">日期：${escHtml(f.today)}</p>`,

  '法定代表人身份证明书': (f) => `
<p style="text-align:center;font-size:15pt;font-weight:bold;margin:0 0 24pt 0;">法定代表人身份证明书</p>
<p style="font-size:12pt;text-indent:2em;line-height:1.8;margin:0 0 8pt 0;"><u>${escHtml(f.legalRep)}</u>，是我单位法定代表人。</p>
<p style="font-size:12pt;text-indent:2em;line-height:1.8;margin:0 0 48pt 0;">特此证明。</p>
<p style="font-size:12pt;text-align:right;margin:0 0 4pt 0;">${escHtml(f.companyName)}</p>
<p style="font-size:12pt;text-align:right;margin:0;">${escHtml(f.today)}</p>`,

  '授权委托书': (f) => `
<p style="text-align:center;font-size:20pt;font-weight:bold;line-height:1.6;margin:0 0 12pt 0;">授权委托书</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;"><strong>委托人：${escHtml(f.companyName)}</strong></p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;">社会统一信用代码：${escHtml(f.creditCode)}</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;">住所：${escHtml(f.address)}</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;">负责人：${escHtml(f.legalRep)}</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;"><strong>受托人：${escHtml(f.handlerName)}</strong>&emsp;&emsp;&emsp;职 务：公司法务</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 4pt 0;">身份证号码：${escHtml(f.handlerIdCard)}&emsp;&emsp;电 话：${escHtml(f.handlerPhone)}</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 12pt 0;">地址：广东省深圳市龙岗区华为坂田基地A区A10</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 6pt 0;">就<u>${escHtml(f.caseName)}一案</u>，委托人现委托${escHtml(f.handlerName)}作为委托人的诉讼代理人。</p>
<p style="font-size:14pt;text-indent:2em;line-height:1.8;margin:0 0 6pt 0;">代理人${escHtml(f.handlerName)}的代理权限为<u>特别代理</u>，包括<u>代为参加诉讼、调查、出庭、参加庭审审理、申请回避、提供证据、质证、进行反驳和辩论，代为承认、放弃、变更诉讼请求，代为进行和解（代为和解金额不超过50万美元）、调解，代为签署、接收相关法律文书等，此授权为特别授权</u>。</p>
<p style="font-size:12pt;margin-left:42pt;text-indent:-18pt;line-height:1.8;margin-bottom:4pt;">本授权不得转授权。</p>
<p style="font-size:12pt;text-indent:2em;line-height:1.8;margin:0 0 24pt 0;">本授权有效期为${escHtml(f.today)}至本案二审程序结束之日。</p>
<p style="font-size:14pt;text-align:right;line-height:1.8;margin:0 0 4pt 0;">委托人（盖章）：<strong>${escHtml(f.companyName)}</strong></p>
<p style="font-size:14pt;text-align:right;line-height:1.8;margin:0;">${escHtml(f.today)}</p>`
};

function resolveDocFields(caseObj) {
  const parties = Storage.getParties();
  const handlers = Storage.getHandlers();

  // Resolve first party for document fields
  const firstMp = (caseObj.myParties && caseObj.myParties[0]) || null;
  let party = null;
  if (firstMp && firstMp.partyId) {
    party = parties.find(p => p.id === firstMp.partyId);
  } else if (!firstMp && caseObj.myPartyId) {
    party = parties.find(p => p.id === caseObj.myPartyId);
  }
  const firstCustom = firstMp ? firstMp.custom : (caseObj.myPartyCustom || '');
  const companyName = party ? party.companyName : firstCustom;
  const address     = party ? (party.address || '') : '';
  const legalRep    = party ? (party.legalRep || '') : '';
  const creditCode  = party ? (party.creditCode || '') : '';

  const handlerObj    = handlers.find(h => h.name === caseObj.handler) || {};
  const handlerName   = handlerObj.name || caseObj.handler || '';
  const handlerIdCard = handlerObj.idCard || '';
  const handlerPhone  = handlerObj.phone || '';

  const now = new Date();
  const today = `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;

  return {
    caseNumber: caseObj.caseNumber || '',
    caseName: caseObj.name || '',
    courtLocation: caseObj.courtLocation || '',
    companyName, address, legalRep, creditCode,
    handlerName, handlerIdCard, handlerPhone, today
  };
}

function openDocModal() {
  const cases = Storage.getCases();
  const sel = document.getElementById('docCaseSelect');
  sel.innerHTML = '<option value="">-- 请选择案件 --</option>' +
    cases.map(c => `<option value="${c.id}">${escHtml(c.name)}</option>`).join('');
  document.querySelectorAll('[name="docType"]').forEach(cb => { cb.checked = false; });
  document.getElementById('docCaseErr').textContent = '';
  document.getElementById('docTypeErr').textContent = '';
  document.getElementById('docModal').style.display = 'flex';
}

function closeDocModal() {
  document.getElementById('docModal').style.display = 'none';
}

function closeDocPreviewModal() {
  document.getElementById('docPreviewModal').style.display = 'none';
}

function getSelectedDocTypes() {
  return [...document.querySelectorAll('[name="docType"]:checked')].map(cb => cb.value);
}

function validateDocForm() {
  let ok = true;
  const caseId = document.getElementById('docCaseSelect').value;
  const selected = getSelectedDocTypes();
  document.getElementById('docCaseErr').textContent = '';
  document.getElementById('docTypeErr').textContent = '';
  if (!caseId) {
    document.getElementById('docCaseErr').textContent = '请选择案件';
    ok = false;
  }
  if (selected.length === 0) {
    document.getElementById('docTypeErr').textContent = '请至少勾选一个文书模版';
    ok = false;
  }
  return ok;
}

function openDocPreview() {
  if (!validateDocForm()) return;
  const caseId = document.getElementById('docCaseSelect').value;
  const selected = getSelectedDocTypes();
  const caseObj = Storage.getCases().find(c => c.id === caseId);
  const fields = resolveDocFields(caseObj);

  const tabsEl = document.getElementById('docTabs');
  const panelsEl = document.getElementById('docPanels');
  tabsEl.innerHTML = '';
  panelsEl.innerHTML = '';

  selected.forEach((docType, i) => {
    const content = DOC_TEMPLATES[docType](fields);
    const tabBtn = document.createElement('button');
    tabBtn.className = 'doc-tab' + (i === 0 ? ' active' : '');
    tabBtn.dataset.idx = i;
    tabBtn.textContent = docType;
    tabsEl.appendChild(tabBtn);

    const panel = document.createElement('div');
    panel.className = 'doc-panel' + (i === 0 ? ' active' : '');
    panel.dataset.idx = i;
    const editDiv = document.createElement('div');
    editDiv.className = 'doc-edit-area';
    editDiv.contentEditable = 'true';
    editDiv.dataset.doctype = docType;
    editDiv.innerHTML = content;
    panel.appendChild(editDiv);
    panelsEl.appendChild(panel);
  });

  tabsEl.querySelectorAll('.doc-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      tabsEl.querySelectorAll('.doc-tab').forEach(t => t.classList.remove('active'));
      panelsEl.querySelectorAll('.doc-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      panelsEl.querySelector(`.doc-panel[data-idx="${tab.dataset.idx}"]`).classList.add('active');
    });
  });

  document.getElementById('docModal').style.display = 'none';
  document.getElementById('docPreviewModal').style.display = 'flex';
}

function exportDocAsWord(filename, htmlContent) {
  const html = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'>
<style>
  body { font-family: "SimSun", "宋体", serif; font-size: 12pt; line-height: 1.8; margin: 2cm; }
  p { margin: 0; padding: 0; }
  u { text-decoration: underline; }
</style>
</head><body>${htmlContent}</body></html>`;
  const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename + '.doc';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function exportAllDocs(fromPreview) {
  if (fromPreview) {
    document.querySelectorAll('.doc-edit-area').forEach(el => {
      exportDocAsWord(el.dataset.doctype, el.innerHTML);
    });
  } else {
    if (!validateDocForm()) return;
    const caseId = document.getElementById('docCaseSelect').value;
    const caseObj = Storage.getCases().find(c => c.id === caseId);
    const fields = resolveDocFields(caseObj);
    getSelectedDocTypes().forEach(docType => {
      exportDocAsWord(docType, DOC_TEMPLATES[docType](fields));
    });
  }
}

// ===== Seed Data =====
function seedParties() {
  if (Storage.getParties().length > 0) return;
  const seed = [
    { companyName: '华为技术有限公司', legalRep: '赵明路', creditCode: '914403001922038216', address: '深圳市龙岗区坂田华为总部办公楼' },
    { companyName: '华为终端有限公司', legalRep: '魏承敏', creditCode: '914419000585344943', address: '广东省东莞市松山湖园区新城路2号' },
    { companyName: '深圳慧通商务有限公司', legalRep: '杜延新', creditCode: '91440300760478875Q', address: '深圳市龙岗区坂田华为基地华为电气科研中心' },
    { companyName: '深圳市海思半导体有限公司', legalRep: '高戟', creditCode: '914403007675804181', address: '深圳市龙岗区坂田华为基地华为电气生产中心' },
    { companyName: '华为数字能源技术有限公司', legalRep: '王辉', creditCode: '91440300MA5GTQ528P', address: '深圳市福田区香蜜湖街道香安社区安托山六路33号安托山总部大厦A座研发39层01号' },
    { companyName: '华为云计算技术有限公司', legalRep: '赵明路', creditCode: '91520900MA6J6CBN9Q', address: '贵州省贵安新区黔中大道交兴功路华为云数据中心' },
    { companyName: '华为技术有限公司北京研究所', legalRep: '姜向中', creditCode: '9111010880200170X7', address: '北京市海淀区木荷路18号1幢等4幢' },
    { companyName: '华为技术有限公司成都研究所', legalRep: '罗卫', creditCode: '9151010072743351XU', address: '成都高新区西源大道1899号' },
    { companyName: '华为技术有限公司武汉研究所', legalRep: '余海波', creditCode: '914201007893316905', address: '武汉东湖高新区高新大道999号武汉未来科技城起步区一期' },
    { companyName: '上海华为技术有限公司', legalRep: '朱立峰', creditCode: '91310000703099764Y', address: '中国上海自由贸易试验区新金桥路2222号' },
    { companyName: '华为技术有限公司杭州研究所', legalRep: '徐峰', creditCode: '91330108727195718R', address: '杭州市滨江区长河街道江虹路410号4幢' },
    { companyName: '华为技术有限公司南京研究所', legalRep: '李峰', creditCode: '91320100716207592F', address: '南京市雨花台区软件大道101号' },
    { companyName: '华为技术有限公司西安研究所', legalRep: '张锦保', creditCode: '91610131726292719Q', address: '陕西省西安市高新区锦业路127号' },
    { companyName: '华为数字技术苏州有限公司', legalRep: '侯金龙', creditCode: '913205945900004996', address: '中国江苏自由贸易试验区苏州片区苏州工业园区江韵路9号' },
    { companyName: '海思技术有限公司', legalRep: '蔡立群', creditCode: '91310118MA1JMHRUXW', address: '上海市青浦区金泽镇培雅南路100号' },
  ].map(p => ({ ...p, id: uuid() }));
  Storage.saveParties(seed);
}

function seedCases() {
  const parties = Storage.getParties();
  const findParty = name => parties.find(p => p.companyName === name);
  const existing = Storage.getCases();
  const existingNames = new Set(existing.map(c => c.name));

  const p2 = findParty('深圳慧通商务有限公司');

  const toAdd = [
    {
      id: uuid(),
      name: '顾可诉苏州华为技术研发有限公司',
      caseNumber: '',
      causeOfAction: '',
      myParties: [{ partyId: null, custom: '苏州华为技术研发有限公司' }],
      oppositeParty: '顾可',
      courtLocation: '苏州',
      courtDate: '2026-03-07',
      evidenceDays: 15,
      stage: '仲裁',
      handler: '王爰越',
      claims: '',
      status: '已结案',
      caseType: '雇员',
      attachments: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    },
    {
      id: uuid(),
      name: '姜波诉深圳慧通商务有限公司、西安景合管理服务有限公司南京分公司',
      caseNumber: '',
      causeOfAction: '',
      myParties: [{ partyId: p2 ? p2.id : null, custom: p2 ? '' : '深圳慧通商务有限公司' }],
      oppositeParty: '姜波',
      courtLocation: '南京',
      courtDate: '2026-04-07',
      evidenceDays: 15,
      stage: '仲裁',
      handler: '王爰越',
      claims: '',
      status: '进行中',
      caseType: '非雇员',
      attachments: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  ].filter(c => !existingNames.has(c.name));

  if (toAdd.length > 0) {
    Storage.saveCases([...toAdd, ...existing]);
  }
}

// ===== Event Bindings =====
document.addEventListener('DOMContentLoaded', () => {
  seedParties();
  seedCases();
  renderCards();

  // Header buttons
  document.getElementById('btnNewCase').addEventListener('click', () => openCaseModal(null));
  document.getElementById('btnHandler').addEventListener('click', openHandlerModal);
  document.getElementById('btnParty').addEventListener('click', openPartyModal);
  document.getElementById('btnBatchImport').addEventListener('click', openBatchModal);

  // Case modal
  document.getElementById('closeCaseModal').addEventListener('click', closeCaseModal);
  document.getElementById('cancelCaseModal').addEventListener('click', closeCaseModal);
  document.getElementById('saveCaseBtn').addEventListener('click', saveCase);
  document.getElementById('deleteCaseBtn').addEventListener('click', () => {
    const id = document.getElementById('caseId').value;
    if (!id) return;
    if (!confirm('确定删除该案件吗？删除后无法恢复。')) return;
    const cases = Storage.getCases().filter(c => c.id !== id);
    Storage.saveCases(cases);
    closeCaseModal();
    renderCards();
  });

  // Evidence hint: update when court date or days change
  document.getElementById('courtDate').addEventListener('change', updateEvidenceHint);
  document.getElementById('evidenceDays').addEventListener('input', updateEvidenceHint);

  // Handler modal
  document.getElementById('closeHandlerModal').addEventListener('click', closeHandlerModal);
  document.getElementById('btnAddHandler').addEventListener('click', () => openHandlerForm(null));
  document.getElementById('saveHandlerBtn').addEventListener('click', saveHandler);
  document.getElementById('cancelHandlerForm').addEventListener('click', showHandlerList);
  document.getElementById('handlerModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeHandlerModal();
  });

  // Party modal
  document.getElementById('closePartyModal').addEventListener('click', closePartyModal);
  document.getElementById('btnAddParty').addEventListener('click', () => openPartyForm(null));
  document.getElementById('savePartyBtn').addEventListener('click', saveParty);
  document.getElementById('cancelPartyForm').addEventListener('click', showPartyList);

  // Add party row
  document.getElementById('btnAddPartyRow').addEventListener('click', () => addPartyRow('', ''));

  // Attachment
  document.getElementById('btnAddAttachment').addEventListener('click', () => {
    document.getElementById('attachmentInput').click();
  });
  document.getElementById('attachmentInput').addEventListener('change', e => {
    Array.from(e.target.files).forEach(f => {
      if (!pendingAttachments.some(a => a.name === f.name)) {
        pendingAttachments.push({ name: f.name });
      }
    });
    e.target.value = '';
    renderAttachmentList();
  });

  // Search
  document.getElementById('searchInput').addEventListener('input', e => {
    currentSearch = e.target.value;
    renderCards();
  });

  // Filter tabs
  document.getElementById('filterTabs').addEventListener('click', e => {
    const tab = e.target.closest('.tab');
    if (!tab) return;
    currentFilter = tab.dataset.status;
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    renderCards();
  });

  // Batch modal
  document.getElementById('closeBatchModal').addEventListener('click', closeBatchModal);
  document.getElementById('cancelBatchModal').addEventListener('click', closeBatchModal);
  document.getElementById('importBatchBtn').addEventListener('click', importBatch);
  document.getElementById('btnAddBatchRow').addEventListener('click', () => {
    const countEl = document.getElementById('addRowCount');
    let count = parseInt(countEl.value) || 1;
    count = Math.min(99, Math.max(1, count));
    for (let i = 0; i < count; i++) addBatchRow();
  });

  // Close modal on overlay click
  document.getElementById('caseModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeCaseModal();
  });
  document.getElementById('partyModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closePartyModal();
  });
  document.getElementById('batchModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeBatchModal();
  });

  // Doc modal
  document.getElementById('btnGenDoc').addEventListener('click', openDocModal);
  document.getElementById('closeDocModal').addEventListener('click', closeDocModal);
  document.getElementById('cancelDocModal').addEventListener('click', closeDocModal);
  document.getElementById('previewDocBtn').addEventListener('click', openDocPreview);
  document.getElementById('exportDocBtn').addEventListener('click', () => exportAllDocs(false));
  document.getElementById('docModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeDocModal();
  });

  // Doc preview modal
  document.getElementById('closeDocPreviewModal').addEventListener('click', closeDocPreviewModal);
  document.getElementById('backToDocModal').addEventListener('click', () => {
    closeDocPreviewModal();
    document.getElementById('docModal').style.display = 'flex';
  });
  document.getElementById('exportFromPreviewBtn').addEventListener('click', () => exportAllDocs(true));
  document.getElementById('docPreviewModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeDocPreviewModal();
  });

  // Keyboard Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      closeCaseModal();
      closeHandlerModal();
      closePartyModal();
      closeBatchModal();
      closeDocModal();
      closeDocPreviewModal();
    }
  });
});
