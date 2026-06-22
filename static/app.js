const state = {
  authToken: localStorage.getItem('bid_parser_token') || '',
  currentUser: null,
  users: [],
  sessionId: '',
  currentProjectId: '',
  sourceFileName: '',
  documentTextExcerpt: '',
  documentTextFull: '',
  packageCount: 0,
  documentSummary: {},
  extractionFields: {},
  registerRows: [],
  documentRegisterRow: {},
  analysis: {},
  matchReview: [],
  followUp: {},
  ourQuotes: [],
  competitorQuotes: [],
  timeline: [],
  registerMode: 'packages',
  templateMeta: null,
  projects: [],
  projectStats: {
    total_projects: 0,
    awarded_projects: 0,
    submitted_projects: 0,
    pending_projects: 0,
    our_quote_rows: 0,
    competitor_quote_rows: 0,
  },
};

const presets = {
  custom: { base_url: '', model: '' },
  openai: { base_url: 'https://api.openai.com/v1', model: 'gpt-4.1-mini' },
  openrouter: { base_url: 'https://openrouter.ai/api/v1', model: 'openai/gpt-4.1-mini' },
};

const followUpFields = [
  { key: 'bid_status', label: '投标状态', type: 'select', options: ['待跟进', '准备投标', '已投标', '放弃', '未投标'] },
  { key: 'award_status', label: '中标状态', type: 'select', options: ['未知', '待定', '已中标', '未中标'] },
  { key: 'award_date', label: '中标日期' },
  { key: 'award_company', label: '中标单位' },
  { key: 'our_award_amount', label: '我司中标价' },
  { key: 'competitor_award_amount', label: '竞对中标价' },
  { key: 'information_source', label: '信息来源' },
  { key: 'register_year', label: '登记年份' },
  { key: 'tracking_note', label: '跟进备注', type: 'textarea' },
];

const ourQuoteColumns = [
  { key: 'package_name', label: '标段/合同包' },
  { key: 'round_no', label: '轮次', type: 'number' },
  { key: 'quote_date', label: '报价日期' },
  { key: 'currency', label: '币种' },
  { key: 'tax_mode', label: '税率/口径' },
  { key: 'unit_price', label: '单价' },
  { key: 'total_price', label: '总价' },
  { key: 'is_submitted', label: '已投递', type: 'checkbox' },
  { key: 'is_awarded', label: '中标', type: 'checkbox' },
  { key: 'remark', label: '备注', type: 'textarea' },
];

const competitorQuoteColumns = [
  { key: 'quote_company', label: '竞对公司' },
  { key: 'package_name', label: '标段/合同包' },
  { key: 'quote_date', label: '报价日期' },
  { key: 'currency', label: '币种' },
  { key: 'unit_price', label: '单价' },
  { key: 'total_price', label: '总价' },
  { key: 'ranking', label: '排名' },
  { key: 'is_awarded', label: '中标', type: 'checkbox' },
  { key: 'source', label: '来源' },
  { key: 'remark', label: '备注', type: 'textarea' },
];

const timelineColumns = [
  { key: 'date', label: '日期' },
  { key: 'type', label: '类型', type: 'select', options: ['parse', 'quote', 'award', 'note'] },
  { key: 'note', label: '内容', type: 'textarea' },
];

const $ = (id) => document.getElementById(id);

async function request(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (state.authToken) {
    headers.set('Authorization', `Bearer ${state.authToken}`);
  }
  const response = await fetch(url, { ...options, headers });
  if (!response.ok) {
    let message = '请求失败';
    try {
      const data = await response.json();
      message = data.detail || message;
    } catch (_) {}
    if (response.status === 401) {
      clearAuth();
      window.location.href = '/login';
    }
    throw new Error(message);
  }
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    return response.json();
  }
  return response.blob();
}

function showToast(message, isError = false) {
  const toast = $('toast');
  toast.textContent = message;
  toast.className = `toast ${isError ? 'error' : ''}`;
  setTimeout(() => {
    toast.className = 'toast hidden';
  }, 2600);
}

function setAuth(token, user) {
  state.authToken = token || '';
  state.currentUser = user || null;
  if (token) {
    localStorage.setItem('bid_parser_token', token);
  } else {
    localStorage.removeItem('bid_parser_token');
  }
  renderCurrentUser();
}

function clearAuth() {
  setAuth('', null);
  state.users = [];
}

function renderCurrentUser() {
  const user = state.currentUser;
  $('userNavBtn').classList.toggle('hidden', user?.role !== 'admin');
}

function renderUserList() {
  if (!state.currentUser || state.currentUser.role !== 'admin') {
    $('userList').innerHTML = `<div class="project-empty">只有管理员可以查看账号列表</div>`;
    return;
  }
  if (!state.users.length) {
    $('userList').innerHTML = `<div class="project-empty">暂无账号</div>`;
    return;
  }
  $('userList').innerHTML = `
    <table class="dense-table">
      <thead><tr><th>用户名</th><th>显示名称</th><th>角色</th><th>创建时间</th></tr></thead>
      <tbody>
        ${state.users
          .map(
            (user) => `
              <tr>
                <td>${user.username || '-'}</td>
                <td>${user.display_name || '-'}</td>
                <td>${user.role === 'admin' ? '管理员' : '普通用户'}</td>
                <td>${user.created_at || '-'}</td>
              </tr>
            `
          )
          .join('')}
      </tbody>
    </table>
  `;
}

function setStatus(id, text, className = 'neutral') {
  const node = $(id);
  node.textContent = text;
  node.className = `pill ${className}`;
}

function switchPanel(targetId) {
  document.querySelectorAll('.nav-link').forEach((node) => {
    node.classList.toggle('active', node.dataset.target === targetId);
  });
  document.querySelectorAll('.content-panel').forEach((node) => {
    node.classList.toggle('hidden', node.id !== targetId);
  });
}

function statusClass(status) {
  if (status === '已明确满足') return 'ok';
  if (status === '疑似风险/否决项') return 'error';
  return 'warn';
}

function normalizeText(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function emptyResult() {
  return {
    document_summary: {
      project_name: '',
      bid_no: '',
      tenderer: '',
      bid_deadline: '',
      open_time: '',
      submission_method: '',
      deposit_amount: '',
      service_period: '',
      qualification_requirements: '',
      vessel_requirements: '',
      technical_business: '',
      quotation: '',
      evaluation_method: '',
    },
    extraction_fields: Object.fromEntries(
      (state.templateMeta?.extraction_fields || []).map((field) => [field.key, { value: '', source_excerpt: '' }])
    ),
    register_rows: [{}],
    document_register_row: {},
    analysis: {
      summary: '',
      qualification_files: [],
      business_points: [],
      scoring_points: [],
      risks: [],
    },
    match_review: [],
    document_text_excerpt: '',
    document_text_full: '',
    package_count: 0,
  };
}

function emptyFollowUp() {
  return {
    bid_status: '待跟进',
    award_status: '未知',
    award_date: '',
    award_company: '',
    our_award_amount: '',
    competitor_award_amount: '',
    tracking_note: '',
    information_source: '',
    register_year: $('sheetName')?.value.trim() || '',
  };
}

function currentRegisterRows() {
  return state.registerMode === 'document' ? [state.documentRegisterRow] : state.registerRows;
}

function collectResultPayload() {
  return {
    document_summary: state.documentSummary,
    extraction_fields: state.extractionFields,
    register_rows: state.registerRows,
    analysis: state.analysis,
    match_review: state.matchReview,
    document_text_excerpt: state.documentTextExcerpt,
    document_text_full: state.documentTextFull,
  };
}

function currentProjectPayload() {
  state.followUp.register_year = $('sheetName').value.trim();
  return {
    title: $('projectTitle').value.trim(),
    source_file_name: state.sourceFileName,
    register_mode: $('registerMode').value,
    sheet_name: $('sheetName').value.trim(),
    result: collectResultPayload(),
    follow_up: state.followUp,
    our_quotes: state.ourQuotes,
    competitor_quotes: state.competitorQuotes,
    timeline: state.timeline,
  };
}

function clearCurrentProject() {
  const result = emptyResult();
  state.currentProjectId = '';
  state.sessionId = '';
  state.sourceFileName = '';
  state.documentTextExcerpt = '';
  state.documentTextFull = '';
  state.packageCount = 0;
  state.documentSummary = result.document_summary;
  state.extractionFields = result.extraction_fields;
  state.registerRows = result.register_rows;
  state.documentRegisterRow = result.document_register_row;
  state.analysis = result.analysis;
  state.matchReview = result.match_review;
  state.followUp = emptyFollowUp();
  state.ourQuotes = [];
  state.competitorQuotes = [];
  state.timeline = [];
  state.registerMode = $('registerMode').value;
  $('projectTitle').value = '';
  $('fileInput').value = '';
  $('chatMessages').innerHTML = '';
  $('summaryPanel').innerHTML = '';
  $('analysisPanel').innerHTML = '';
  $('reviewPanel').innerHTML = '';
  $('followUpForm').innerHTML = '';
  $('ourQuotesTable').innerHTML = '';
  $('competitorQuotesTable').innerHTML = '';
  $('timelineTable').innerHTML = '';
  $('extractionTable').innerHTML = '';
  $('registerTable').innerHTML = '';
  $('resultSection').classList.add('hidden');
  $('resultEmpty').classList.remove('hidden');
  setStatus('fileStatus', '待解析');
  renderProjectList();
}

function renderSummary(summary) {
  const items = [
    ['项目名称', summary.project_name],
    ['招标编号', summary.bid_no],
    ['招标人', summary.tenderer],
    ['投标截止时间', summary.bid_deadline],
    ['开标时间', summary.open_time],
    ['保证金', summary.deposit_amount],
    ['服务期限', summary.service_period],
    ['递交方式', summary.submission_method],
    ['识别标的数', state.packageCount || 1],
    ['投标状态', state.followUp.bid_status || '待跟进'],
    ['中标状态', state.followUp.award_status || '未知'],
    ['我司报价条数', state.ourQuotes.length],
  ];
  $('summaryPanel').innerHTML = items
    .map(([label, value]) => `<article class="summary-card"><span>${label}</span><strong>${value || '-'}</strong></article>`)
    .join('');
}

function renderAnalysis(analysis) {
  const blocks = [
    ['解析结论', [analysis.summary]],
    ['资格文件清单', analysis.qualification_files],
    ['报价与商务关注点', analysis.business_points],
    ['评标得分关注点', analysis.scoring_points],
    ['风险提示', analysis.risks],
  ];
  $('analysisPanel').innerHTML = blocks
    .map(([title, items]) => {
      const lines = (items || []).filter(Boolean);
      const content = lines.length ? lines.map((item) => `<li>${item}</li>`).join('') : '<li>-</li>';
      return `<article class="analysis-card"><h4>${title}</h4><ul>${content}</ul></article>`;
    })
    .join('');
}

function renderReview(items) {
  if (!items?.length) {
    $('reviewPanel').innerHTML = `<article class="review-item"><p>当前没有额外的匹配复核项。</p></article>`;
    return;
  }
  $('reviewPanel').innerHTML = items
    .map(
      (item) => `
        <article class="review-item">
          <div class="review-head">
            <strong>${item.area || '待分类'}</strong>
            <span class="status ${statusClass(item.status)}">${item.status || '需人工确认'}</span>
          </div>
          <p><b>事项：</b>${item.item || '-'}</p>
          <p><b>原因：</b>${item.reason || '-'}</p>
          <p><b>建议动作：</b>${item.action || '-'}</p>
        </article>
      `
    )
    .join('');
}

function renderExtraction() {
  const rows = state.templateMeta?.extraction_fields || [];
  $('extractionTable').innerHTML = `
    <table>
      <thead><tr><th>字段</th><th>摘取结果</th><th>原文依据</th></tr></thead>
      <tbody>
        ${rows
          .map((field) => {
            const value = state.extractionFields[field.key] || { value: '', source_excerpt: '' };
            return `
              <tr>
                <td class="label-col">${field.label}</td>
                <td><textarea data-type="extraction-value" data-key="${field.key}" rows="4">${value.value || ''}</textarea></td>
                <td><textarea data-type="extraction-source" data-key="${field.key}" rows="6">${value.source_excerpt || ''}</textarea></td>
              </tr>
            `;
          })
          .join('')}
      </tbody>
    </table>
  `;
}

function renderRegister() {
  const rows = currentRegisterRows();
  const columns = state.templateMeta?.register_columns || [];
  $('registerTable').innerHTML = `
    <table>
      <thead><tr>${columns.map((column) => `<th>${column.label}</th>`).join('')}</tr></thead>
      <tbody>
        ${rows
          .map(
            (row, rowIndex) => `
              <tr>
                ${columns
                  .map(
                    (column) => `
                      <td><textarea data-type="register" data-row="${rowIndex}" data-key="${column.key}" rows="3">${row[column.key] || ''}</textarea></td>
                    `
                  )
                  .join('')}
              </tr>
            `
          )
          .join('')}
      </tbody>
    </table>
  `;
}

function renderFollowUp() {
  $('followUpForm').innerHTML = followUpFields
    .map((field) => {
      const value = state.followUp[field.key] || '';
      if (field.type === 'select') {
        return `
          <label class="field">
            <span>${field.label}</span>
            <select data-type="follow-up" data-key="${field.key}">
              ${field.options
                .map((option) => `<option value="${option}" ${option === value ? 'selected' : ''}>${option}</option>`)
                .join('')}
            </select>
          </label>
        `;
      }
      if (field.type === 'textarea') {
        return `
          <label class="field field-span-2">
            <span>${field.label}</span>
            <textarea data-type="follow-up" data-key="${field.key}" rows="4">${value}</textarea>
          </label>
        `;
      }
      return `
        <label class="field">
          <span>${field.label}</span>
          <input data-type="follow-up" data-key="${field.key}" value="${value}" />
        </label>
      `;
    })
    .join('');
}

function tableCellMarkup(kind, rowIndex, column, value) {
  if (column.type === 'checkbox') {
    return `<input type="checkbox" data-type="${kind}" data-row="${rowIndex}" data-key="${column.key}" ${value ? 'checked' : ''} />`;
  }
  if (column.type === 'select') {
    return `
      <select data-type="${kind}" data-row="${rowIndex}" data-key="${column.key}">
        ${column.options
          .map((option) => `<option value="${option}" ${option === value ? 'selected' : ''}>${option}</option>`)
          .join('')}
      </select>
    `;
  }
  if (column.type === 'textarea') {
    return `<textarea data-type="${kind}" data-row="${rowIndex}" data-key="${column.key}" rows="3">${value || ''}</textarea>`;
  }
  return `<input data-type="${kind}" data-row="${rowIndex}" data-key="${column.key}" value="${value || ''}" ${column.type === 'number' ? 'type="number" min="1"' : ''} />`;
}

function renderEditableTable(containerId, kind, rows, columns, emptyText) {
  if (!rows.length) {
    $(containerId).innerHTML = `<div class="project-empty">${emptyText}</div>`;
    return;
  }
  $(containerId).innerHTML = `
    <table class="dense-table">
      <thead>
        <tr>
          ${columns.map((column) => `<th>${column.label}</th>`).join('')}
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row, rowIndex) => `
              <tr>
                ${columns.map((column) => `<td>${tableCellMarkup(kind, rowIndex, column, row[column.key])}</td>`).join('')}
                <td class="action-col"><button data-action="remove-row" data-kind="${kind}" data-row="${rowIndex}" class="ghost danger">删除</button></td>
              </tr>
            `
          )
          .join('')}
      </tbody>
    </table>
  `;
}

function renderTracking() {
  renderFollowUp();
  renderEditableTable('ourQuotesTable', 'our-quote', state.ourQuotes, ourQuoteColumns, '还没有录入我司报价');
  renderEditableTable('competitorQuotesTable', 'competitor-quote', state.competitorQuotes, competitorQuoteColumns, '还没有录入竞对报价');
  renderEditableTable('timelineTable', 'timeline', state.timeline, timelineColumns, '还没有跟进记录');
}

function bindDynamicEditors() {
  document.querySelectorAll("textarea[data-type='extraction-value']").forEach((node) => {
    node.oninput = (event) => {
      const key = event.target.dataset.key;
      state.extractionFields[key].value = event.target.value;
    };
  });
  document.querySelectorAll("textarea[data-type='extraction-source']").forEach((node) => {
    node.oninput = (event) => {
      const key = event.target.dataset.key;
      state.extractionFields[key].source_excerpt = event.target.value;
    };
  });
  document.querySelectorAll("textarea[data-type='register']").forEach((node) => {
    node.oninput = (event) => {
      const rowIndex = Number(event.target.dataset.row);
      const key = event.target.dataset.key;
      const rows = currentRegisterRows();
      rows[rowIndex][key] = event.target.value;
      if (state.registerMode === 'document') {
        state.documentRegisterRow = rows[0];
      } else {
        state.registerRows = rows;
      }
    };
  });
  document.querySelectorAll("[data-type='follow-up']").forEach((node) => {
    const handler = (event) => {
      const key = event.target.dataset.key;
      state.followUp[key] = event.target.value;
      if (key === 'register_year' && !$('sheetName').value.trim()) {
        $('sheetName').value = event.target.value;
      }
      renderSummary(state.documentSummary);
    };
    node.oninput = handler;
    node.onchange = handler;
  });
  document.querySelectorAll("[data-type='our-quote'], [data-type='competitor-quote'], [data-type='timeline']").forEach((node) => {
    const handler = (event) => {
      const kind = event.target.dataset.type;
      const rowIndex = Number(event.target.dataset.row);
      const key = event.target.dataset.key;
      const target =
        kind === 'our-quote' ? state.ourQuotes : kind === 'competitor-quote' ? state.competitorQuotes : state.timeline;
      target[rowIndex][key] = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
      renderSummary(state.documentSummary);
    };
    node.oninput = handler;
    node.onchange = handler;
  });
  document.querySelectorAll("[data-action='remove-row']").forEach((node) => {
    node.onclick = () => {
      const rowIndex = Number(node.dataset.row);
      const kind = node.dataset.kind;
      const target =
        kind === 'our-quote' ? state.ourQuotes : kind === 'competitor-quote' ? state.competitorQuotes : state.timeline;
      target.splice(rowIndex, 1);
      renderTracking();
      bindDynamicEditors();
      renderSummary(state.documentSummary);
    };
  });
}

function appendChat(role, text) {
  const item = document.createElement('div');
  item.className = `chat-message ${role}`;
  item.innerHTML = `<span>${role === 'user' ? '你' : 'AI'}</span><p>${normalizeText(text).replace(/\n/g, '<br>')}</p>`;
  $('chatMessages').appendChild(item);
  $('chatMessages').scrollTop = $('chatMessages').scrollHeight;
}

function applyResult(result) {
  state.documentSummary = result.document_summary;
  state.extractionFields = result.extraction_fields;
  state.registerRows = result.register_rows;
  state.documentRegisterRow = result.document_register_row;
  state.analysis = result.analysis;
  state.matchReview = result.match_review;
  state.documentTextExcerpt = result.document_text_excerpt || '';
  state.documentTextFull = result.document_text_full || result.document_text_excerpt || '';
  state.packageCount = Number(result.package_count || (result.register_rows || []).length || 0);
  renderSummary(result.document_summary);
  renderAnalysis(result.analysis);
  renderReview(result.match_review);
  renderTracking();
  renderExtraction();
  renderRegister();
  bindDynamicEditors();
  $('resultEmpty').classList.add('hidden');
  $('resultSection').classList.remove('hidden');
}

function renderProjectList() {
  $('projectCount').textContent = `${state.projects.length} 条`;
  $('historyStats').innerHTML = [
    ['项目总数', state.projectStats.total_projects],
    ['已中标', state.projectStats.awarded_projects],
    ['已投标', state.projectStats.submitted_projects],
    ['待跟进', state.projectStats.pending_projects],
    ['我司报价条数', state.projectStats.our_quote_rows],
    ['竞对报价条数', state.projectStats.competitor_quote_rows],
  ]
    .map(
      ([label, value]) => `
        <article class="history-stat">
          <span>${label}</span>
          <strong>${value || 0}</strong>
        </article>
      `
    )
    .join('');
  if (!state.projects.length) {
    $('projectList').innerHTML = `<div class="project-empty">没有匹配的历史记录</div>`;
    return;
  }
  $('projectList').innerHTML = state.projects
    .map(
      (project) => `
        <button class="project-item ${project.project_id === state.currentProjectId ? 'active' : ''}" data-project-id="${project.project_id}">
          <div class="project-card-head">
            <strong>${project.title || '未命名项目'}</strong>
            <span class="status-chip ${project.award_status === '已中标' ? 'ok' : project.award_status === '未中标' ? 'error' : 'warn'}">${project.award_status || '未知'}</span>
          </div>
          <span>${project.bid_no || project.project_name || project.source_file_name || '暂无编号'}</span>
          <div class="project-tags">
            <span>${project.tenderer || '暂无招标人'}</span>
            <span>${project.register_year || '未填年份'}</span>
            <span>${project.bid_status || '待跟进'}</span>
          </div>
          <small>${(project.updated_at || '').replace('T', ' ')}</small>
        </button>
      `
    )
    .join('');
  document.querySelectorAll('.project-item').forEach((node) => {
    node.onclick = async () => {
      try {
        await loadProject(node.dataset.projectId);
      } catch (error) {
        showToast(error.message, true);
      }
    };
  });
}

async function loadProjects() {
  const params = new URLSearchParams(currentHistoryFilters());
  const result = await request(`/api/projects?${params.toString()}`);
  state.projects = result.items || [];
  state.projectStats = result.stats || state.projectStats;
  renderProjectList();
}

async function loadCurrentUser() {
  const result = await request('/api/auth/me');
  setAuth(state.authToken, result.user);
}

async function login() {
  const result = await request('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      username: $('loginUsername').value.trim(),
      password: $('loginPassword').value,
    }),
  });
  setAuth(result.token, result.user);
  $('loginPassword').value = '';
}

async function logout() {
  try {
    await request('/api/auth/logout', { method: 'POST' });
  } catch (_) {}
  clearAuth();
  window.location.href = '/login';
}

async function loadUsers() {
  if (state.currentUser?.role !== 'admin') {
    state.users = [];
    renderUserList();
    return;
  }
  const result = await request('/api/users');
  state.users = result.items || [];
  renderUserList();
}

async function createUser() {
  await request('/api/users', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      username: $('newUsername').value.trim(),
      display_name: $('newDisplayName').value.trim(),
      password: $('newPassword').value,
      role: $('newUserRole').value,
    }),
  });
  $('newUsername').value = '';
  $('newDisplayName').value = '';
  $('newPassword').value = '';
  $('newUserRole').value = 'user';
  await loadUsers();
  showToast('账号创建成功');
}

async function loadMeta() {
  const meta = await request('/api/template-meta');
  state.templateMeta = meta;
  $('sheetName').value = meta.default_sheet || '';
}

async function loadConfig() {
  const config = await request('/api/config');
  $('baseUrl').value = config.base_url || '';
  $('modelName').value = config.model || '';
  $('apiKey').value = '';
  $('temperature').value = config.temperature ?? 0.1;
  setStatus('configStatus', config.base_url && config.has_api_key ? '已保存' : '未保存', config.base_url && config.has_api_key ? 'ok' : 'neutral');
}

function readConfigForm() {
  return {
    base_url: $('baseUrl').value.trim(),
    api_key: $('apiKey').value.trim(),
    model: $('modelName').value.trim(),
    temperature: Number($('temperature').value || 0.1),
  };
}

function applyPreset(name) {
  const preset = presets[name];
  if (!preset || name === 'custom') return;
  $('baseUrl').value = preset.base_url;
  if (!$('modelName').value.trim()) {
    $('modelName').value = preset.model;
  }
}

async function saveConfig() {
  await request('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(readConfigForm()),
  });
  setStatus('configStatus', '已保存', 'ok');
  showToast('AI 配置已保存');
}

async function testConfig() {
  const result = await request('/api/config/test', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(readConfigForm()),
  });
  showToast(`连通成功：${result.message}`);
}

function ensureDefaultTracking() {
  if (!state.timeline.length) {
    state.timeline = [
      {
        id: crypto.randomUUID ? crypto.randomUUID().slice(0, 12) : String(Date.now()),
        date: new Date().toISOString().slice(0, 10),
        type: 'parse',
        note: '完成标书解析',
      },
    ];
  }
}

async function parseFile() {
  const file = $('fileInput').files[0];
  if (!file) throw new Error('请先选择标书文件');
  state.currentProjectId = '';
  state.registerMode = $('registerMode').value;
  setStatus('fileStatus', '解析中', 'warn');
  const formData = new FormData();
  formData.append('file', file);
  const result = await request('/api/parse', { method: 'POST', body: formData });
  state.sessionId = result.session_id;
  state.sourceFileName = result.source_file_name || file.name;
  state.followUp = emptyFollowUp();
  state.followUp.register_year = $('sheetName').value.trim();
  state.ourQuotes = [];
  state.competitorQuotes = [];
  state.timeline = [];
  ensureDefaultTracking();
  if (!$('projectTitle').value.trim()) {
    $('projectTitle').value = result.document_summary.project_name || file.name;
  }
  $('chatMessages').innerHTML = '';
  appendChat('assistant', result.analysis.summary || '解析完成，可以继续追问。');
  applyResult(result);
  setStatus('fileStatus', '解析完成', 'ok');
  showToast('解析完成，现在可以保存到历史记录');
  switchPanel('overviewPanel');
}

function hasProjectContent() {
  return !!($('projectTitle').value.trim() || state.documentSummary.project_name || state.sourceFileName);
}

async function saveCurrentProject() {
  if (!hasProjectContent()) throw new Error('当前没有可保存的内容');
  const payload = currentProjectPayload();
  if (state.currentProjectId) {
    await request(`/api/projects/${state.currentProjectId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    showToast('记录已更新');
  } else {
    const created = await request('/api/projects', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    state.currentProjectId = created.project_id;
    showToast('记录已保存到历史列表');
  }
  await loadProjects();
}

async function loadProject(projectId) {
  const project = await request(`/api/projects/${projectId}`);
  state.currentProjectId = project.project_id;
  state.sessionId = project.session_id;
  state.sourceFileName = project.source_file_name || '';
  state.registerMode = project.register_mode || 'packages';
  state.followUp = project.follow_up || emptyFollowUp();
  state.ourQuotes = project.our_quotes || [];
  state.competitorQuotes = project.competitor_quotes || [];
  state.timeline = project.timeline || [];
  $('registerMode').value = state.registerMode;
  $('sheetName').value = project.sheet_name || state.followUp.register_year || $('sheetName').value;
  $('projectTitle').value = project.title || '';
  $('chatMessages').innerHTML = '';
  applyResult(project.result);
  setStatus('fileStatus', '已加载历史记录', 'ok');
  renderProjectList();
  switchPanel('historyPanel');
}

async function deleteCurrentProject() {
  if (!state.currentProjectId) throw new Error('当前没有已保存的历史记录');
  const name = $('projectTitle').value.trim() || '当前记录';
  if (!window.confirm(`确认删除“${name}”吗？删除后无法恢复。`)) return;
  await request(`/api/projects/${state.currentProjectId}`, { method: 'DELETE' });
  clearCurrentProject();
  await loadProjects();
  showToast('记录已删除');
}

async function sendChat(question) {
  if (!state.sessionId) throw new Error('请先解析标书或打开一条历史记录');
  appendChat('user', question);
  const result = await request('/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ session_id: state.sessionId, question }),
  });
  appendChat('assistant', result.answer);
}

async function downloadFile(url, payload, filename) {
  const blob = await request(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function currentHistoryFilters() {
  return {
    q: $('projectSearch').value.trim(),
    year: $('historyYear').value.trim(),
    bid_status: $('historyBidStatus').value,
    award_status: $('historyAwardStatus').value,
  };
}

function newQuote(defaultCompany = '') {
  return {
    id: crypto.randomUUID ? crypto.randomUUID().slice(0, 12) : String(Date.now()),
    package_name: '',
    round_no: 1,
    quote_date: '',
    quote_company: defaultCompany,
    currency: 'CNY',
    tax_mode: '',
    unit_price: '',
    total_price: '',
    ranking: '',
    is_submitted: false,
    is_awarded: false,
    source: '',
    remark: '',
  };
}

function newTimeline() {
  return {
    id: crypto.randomUUID ? crypto.randomUUID().slice(0, 12) : String(Date.now()),
    date: new Date().toISOString().slice(0, 10),
    type: 'note',
    note: '',
  };
}

function bindEvents() {
  document.querySelectorAll('.nav-link').forEach((node) => {
    node.onclick = () => switchPanel(node.dataset.target);
  });

  ['projectSearch', 'historyYear', 'historyBidStatus', 'historyAwardStatus'].forEach((id) => {
    $(id).oninput = async () => {
      try {
        await loadProjects();
      } catch (error) {
        showToast(error.message, true);
      }
    };
    $(id).onchange = $(id).oninput;
  });

  $('configPreset').onchange = (event) => applyPreset(event.target.value);

  $('logoutBtn').onclick = async () => {
    await logout();
  };

  $('saveConfigBtn').onclick = async () => {
    try {
      await saveConfig();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('testConfigBtn').onclick = async () => {
    try {
      await testConfig();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('parseBtn').onclick = async () => {
    try {
      await parseFile();
    } catch (error) {
      setStatus('fileStatus', '解析失败', 'error');
      showToast(error.message, true);
    }
  };

  $('saveProjectBtn').onclick = async () => {
    try {
      await saveCurrentProject();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('deleteProjectBtn').onclick = async () => {
    try {
      await deleteCurrentProject();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('registerMode').onchange = () => {
    state.registerMode = $('registerMode').value;
    if (!$('resultSection').classList.contains('hidden')) {
      renderRegister();
      bindDynamicEditors();
    }
  };

  $('sheetName').oninput = () => {
    state.followUp.register_year = $('sheetName').value.trim();
  };

  $('addOurQuoteBtn').onclick = () => {
    state.ourQuotes.push(newQuote('我司'));
    renderTracking();
    bindDynamicEditors();
    renderSummary(state.documentSummary);
  };

  $('addCompetitorQuoteBtn').onclick = () => {
    state.competitorQuotes.push(newQuote(''));
    renderTracking();
    bindDynamicEditors();
    renderSummary(state.documentSummary);
  };

  $('addTimelineBtn').onclick = () => {
    state.timeline.push(newTimeline());
    renderTracking();
    bindDynamicEditors();
  };

  $('sendChatBtn').onclick = async () => {
    const question = $('chatInput').value.trim();
    if (!question) return;
    $('chatInput').value = '';
    try {
      await sendChat(question);
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('askDefaultBtn').onclick = async () => {
    try {
      await sendChat('请重新总结这份标书的关键节点、资格要求、报价要求、评标标准和主要风险。');
      switchPanel('chatPanel');
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportOverviewBtn').onclick = async () => {
    try {
      await downloadFile(
        '/api/export/overview',
        {
          title: $('projectTitle').value.trim(),
          result: collectResultPayload(),
          follow_up: state.followUp,
          our_quotes: state.ourQuotes,
          competitor_quotes: state.competitorQuotes,
          timeline: state.timeline,
        },
        '解析总览.xlsx'
      );
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportExtractionBtn').onclick = async () => {
    try {
      await downloadFile('/api/export/extraction', { result: collectResultPayload() }, '标书摘取.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportRegisterBtn').onclick = async () => {
    try {
      await downloadFile('/api/export/register', { rows: currentRegisterRows(), sheet_name: $('sheetName').value }, '招标登记.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportOurQuotesBtn').onclick = async () => {
    try {
      await downloadFile(
        '/api/export/our-quotes',
        {
          title: $('projectTitle').value.trim(),
          follow_up: state.followUp,
          rows: state.ourQuotes,
        },
        '我司报价一览表.xlsx'
      );
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportCompetitorQuotesBtn').onclick = async () => {
    try {
      await downloadFile(
        '/api/export/competitor-quotes',
        {
          title: $('projectTitle').value.trim(),
          follow_up: state.followUp,
          rows: state.competitorQuotes,
        },
        '竞对报价对比.xlsx'
      );
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('exportLedgerBtn').onclick = async () => {
    try {
      await downloadFile('/api/export/ledger', currentHistoryFilters(), '历史台账汇总.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('createUserBtn').onclick = async () => {
    try {
      await createUser();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('refreshUsersBtn').onclick = async () => {
    try {
      await loadUsers();
    } catch (error) {
      showToast(error.message, true);
    }
  };
}

async function init() {
  try {
    bindEvents();
    renderCurrentUser();
    renderUserList();
    if (!state.authToken) {
      window.location.href = '/login';
      return;
    }
    await loadCurrentUser();
    await loadMeta();
    await loadConfig();
    clearCurrentProject();
    await loadProjects();
    if (state.currentUser?.role === 'admin') {
      await loadUsers();
    }
    switchPanel('overviewPanel');
  } catch (error) {
    if (!state.authToken) {
      window.location.href = '/login';
      return;
    }
    showToast(error.message, true);
  }
}

init();
