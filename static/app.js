const state = {
  authToken: localStorage.getItem('bid_parser_token') || '',
  currentUser: null,
  users: [],
  sessionId: '',
  currentProjectId: '',
  sourceFileName: '',
  projectSource: {},
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
  marketKind: 'cargo',
  marketCurrentId: '',
  marketRecord: {},
  marketItems: [],
  marketStats: {},
  marketRequestSeq: 0,
  marketReport: null,
  agentConversationId: localStorage.getItem('shipping_agent_conversation_id') || '',
  agentTranscript: [],
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

const marketCargoFields = [
  { key: 'board_type', label: '货盘类型', type: 'select', options: ['长期货盘', '即时货盘'] },
  { key: 'segment', label: '业务板块', type: 'select', options: ['内贸化', '内贸油', '外贸'] },
  { key: 'cargo_name', label: '货品名称' },
  { key: 'tonnage', label: '货物吨数' },
  { key: 'load_port', label: '装港' },
  { key: 'discharge_port', label: '卸港' },
  { key: 'laycan', label: '装载期' },
  { key: 'cargo_owner', label: '货主/租家' },
  { key: 'cargo_date', label: '货盘日期' },
  { key: 'source', label: '货盘来源' },
  { key: 'status', label: '是否达成合作', type: 'select', options: ['跟进中', '已成交', '未成交', '放弃'] },
  { key: 'final_price', label: '成交价' },
  { key: 'deal_date', label: '成交日期', type: 'date' },
  { key: 'deal_price', label: '最终成交价格' },
  { key: 'price_unit', label: '价格单位' },
  { key: 'currency', label: '币种', type: 'select', options: ['CNY', 'USD'] },
  { key: 'route', label: '航线' },
  { key: 'cargo_standard_name', label: '货品标准名' },
  { key: 'executing_vessel', label: '最终执行船舶' },
  { key: 'market_info', label: '市场了解信息', type: 'textarea' },
  { key: 'loss_reason', label: '未达成/放弃原因', type: 'textarea' },
  { key: 'competitor_name', label: '竞争对手' },
  { key: 'competitor_price', label: '竞争对手价格' },
  { key: 'remark', label: '备注', type: 'textarea' },
];

const marketNewbuildingFields = [
  { key: 'stage', label: '建造阶段', type: 'select', options: ['已完造并出厂投运', '合同签订未开造', '预计2027年完造出厂', '预计2028年完造出厂', '信息待获取'] },
  { key: 'ship_name', label: '船名' },
  { key: 'update_date', label: '更新日期' },
  { key: 'shipyard', label: '造船厂' },
  { key: 'owner', label: '船东' },
  { key: 'dwt', label: '载重吨DWT' },
  { key: 'build_status', label: '当前建造状态' },
  { key: 'delivery_time', label: '预计交付时间' },
  { key: 'actual_delivery_date', label: '实际出厂时间', type: 'date' },
  { key: 'status_update_date', label: '状态更新时间', type: 'date' },
  { key: 'status_note', label: '状态变化备注', type: 'textarea' },
  { key: 'contract_date', label: '合同签订日期' },
  { key: 'contract_price', label: '新造船合同价格' },
  { key: 'ship_type', label: '船型' },
  { key: 'source', label: '信息来源' },
  { key: 'remark', label: '备注', type: 'textarea' },
];

const $ = (id) => document.getElementById(id);

function bindClick(id, handler) {
  const node = $(id);
  if (node) {
    node.onclick = handler;
  }
  return node;
}

function bindChange(id, handler) {
  const node = $(id);
  if (node) {
    node.onchange = handler;
  }
  return node;
}

function bindInput(id, handler, syncChange = false) {
  const node = $(id);
  if (node) {
    node.oninput = handler;
    if (syncChange) {
      node.onchange = handler;
    }
  }
  return node;
}

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
  const userNavBtn = $('userNavBtn');
  if (userNavBtn) {
    userNavBtn.classList.toggle('hidden', user?.role !== 'admin');
  }
  const userPanel = $('userPanel');
  if (userPanel) {
    userPanel.classList.toggle('hidden', user?.role !== 'admin');
  }
}

function renderUserList() {
  const userList = $('userList');
  if (!userList) return;
  if (!state.currentUser || state.currentUser.role !== 'admin') {
    userList.innerHTML = `<div class="project-empty">只有管理员可以查看账号列表</div>`;
    return;
  }
  if (!state.users.length) {
    userList.innerHTML = `<div class="project-empty">暂无账号</div>`;
    return;
  }
  userList.innerHTML = `
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

function switchWorkspace(workspaceId) {
  document.querySelectorAll('.main-tab').forEach((node) => {
    node.classList.toggle('active', node.dataset.workspace === workspaceId);
  });
  document.querySelectorAll('.business-workspace').forEach((node) => {
    node.classList.toggle('active', node.id === workspaceId);
  });
  if (workspaceId === 'marketWorkspace') {
    const marketPanel = $('marketSkillPanel');
    if (marketPanel) marketPanel.classList.remove('hidden');
  }
}

function switchPanel(targetId) {
  switchWorkspace('bidWorkspace');
  document.querySelectorAll('.nav-link').forEach((node) => {
    node.classList.toggle('active', node.dataset.target === targetId);
  });
  document.querySelectorAll('#bidWorkspace .content-panel').forEach((node) => {
    node.classList.toggle('hidden', node.id !== targetId);
  });
}

function toggleSettings(forceOpen = null) {
  const panel = $('settingsPanel');
  if (!panel) return;
  const shouldOpen = forceOpen === null ? panel.classList.contains('hidden') : forceOpen;
  panel.classList.toggle('hidden', !shouldOpen);
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

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
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
    source_channel: state.projectSource.source_channel || '',
    source_open_id: state.projectSource.source_open_id || '',
    source_chat_id: state.projectSource.source_chat_id || '',
    source_message_id: state.projectSource.source_message_id || '',
    source_text: state.projectSource.source_text || '',
    confirmed_at: state.projectSource.confirmed_at || '',
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
  state.projectSource = {};
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
  renderProjectSourceSummary();
}

function renderProjectSourceSummary() {
  const source = state.projectSource || {};
  const isFeishu = source.source_channel === 'feishu';
  if (!isFeishu && !source.source_text) return;
  const meta = [
    isFeishu ? '飞书录入' : source.source_channel || '来源记录',
    source.confirmed_at ? `确认：${String(source.confirmed_at).replace('T', ' ')}` : '',
    source.source_message_id ? `消息：${source.source_message_id}` : '',
  ].filter(Boolean);
  $('summaryPanel').insertAdjacentHTML(
    'beforeend',
    `
      <article class="summary-card project-source-summary">
        <span>数据来源</span>
        <strong>${meta.map((item) => `<em>${escapeHtml(item)}</em>`).join('')}</strong>
        ${source.source_text ? `<small>${escapeHtml(source.source_text)}</small>` : ''}
      </article>
    `
  );
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

function agentPageContext() {
  const workspace = document.querySelector('.main-tab.active')?.dataset.workspace === 'marketWorkspace' ? 'market' : 'bid';
  const panel = document.querySelector('#bidWorkspace .nav-link.active')?.dataset.target || '';
  return {
    workspace,
    panel,
    market_kind: state.marketKind,
    current_market_record_id: state.marketCurrentId,
    market_record: state.marketRecord,
    filters: workspace === 'market' ? marketFilters() : currentHistoryFilters(),
    current_project_id: state.currentProjectId,
    project_form: {
      title: $('projectTitle')?.value.trim() || '',
      follow_up: state.followUp,
      our_quotes: state.ourQuotes,
      competitor_quotes: state.competitorQuotes,
      timeline: state.timeline,
    },
    session_id: state.sessionId,
  };
}

function renderAgentMessages() {
  const node = $('agentMessages');
  if (!node) return;
  node.innerHTML = (state.agentTranscript || [])
    .map((item) => `<div class="chat-message ${item.role === 'user' ? 'user' : ''}"><span>${item.role === 'user' ? '你' : 'AI'}</span><p>${escapeHtml(normalizeText(item.content)).replace(/\n/g, '<br>')}</p></div>`)
    .join('');
  node.scrollTop = node.scrollHeight;
}

function setAgentPending(pending = {}) {
  const node = $('agentPending');
  if (!node) return;
  const changes = pending.changes || {};
  const rows = Object.entries(changes).map(([key, value]) => `${key}：${value.before || '空'} → ${value.after || '空'}`);
  node.classList.toggle('hidden', !pending.type);
  node.textContent = pending.type ? `待确认变更${rows.length ? `：${rows.join('；')}` : ''}。回复“确认保存”或“取消”。` : '';
}

function toggleAgentDrawer(forceOpen = null) {
  const drawer = $('agentDrawer');
  if (!drawer) return;
  const open = forceOpen === null ? drawer.classList.contains('hidden') : forceOpen;
  drawer.classList.toggle('hidden', !open);
  if (open) {
    renderAgentMessages();
    $('agentInput')?.focus();
  }
}

async function loadAgentSession() {
  if (!state.agentConversationId) return;
  try {
    const result = await request(`/api/agent/session/${state.agentConversationId}`);
    state.agentTranscript = result.transcript || [];
    renderAgentMessages();
    setAgentPending(result.pending_change || {});
  } catch (error) {
    state.agentConversationId = '';
    localStorage.removeItem('shipping_agent_conversation_id');
  }
}

async function applyAgentActions(actions = []) {
  for (const action of actions) {
    if (!action || !action.type) continue;
    if (action.type === 'navigate') {
      if (action.workspace === 'market') {
        switchWorkspace('marketWorkspace');
        if (action.kind && action.kind !== state.marketKind) {
          state.marketKind = action.kind;
          setMarketRecord();
          await loadMarketRecords();
        }
      } else {
        switchWorkspace('bidWorkspace');
        if (action.panel) switchPanel(action.panel);
      }
    } else if (action.type === 'load_market_record' || action.type === 'patch_market_form') {
      switchWorkspace('marketWorkspace');
      if (action.kind && action.kind !== state.marketKind) state.marketKind = action.kind;
      if (action.record) {
        setMarketRecord(action.record, { persisted: action.persisted !== false });
      } else if (action.record_id) {
        const result = await request(`/api/market-skill/${state.marketKind}/${action.record_id}`);
        setMarketRecord(result.record, { persisted: true });
      }
      await loadMarketRecords();
    } else if (action.type === 'load_project') {
      await loadProject(action.project_id);
    } else if (action.type === 'patch_project_form') {
      switchWorkspace('bidWorkspace');
      switchPanel('trackingPanel');
      const project = action.project || {};
      state.followUp = project.follow_up || state.followUp;
      state.ourQuotes = project.our_quotes || state.ourQuotes;
      state.competitorQuotes = project.competitor_quotes || state.competitorQuotes;
      state.timeline = project.timeline || state.timeline;
      renderTracking();
      bindDynamicEditors();
    } else if (action.type === 'render_market_report') {
      switchWorkspace('marketWorkspace');
      state.marketKind = action.kind || state.marketKind;
      state.marketReport = action.report || null;
      renderMarketList();
    } else if (action.type === 'apply_bid_result') {
      switchWorkspace('bidWorkspace');
      if (action.session_id) state.sessionId = action.session_id;
      if (action.result) {
        applyResult(action.result);
        switchPanel('overviewPanel');
      }
    } else if (action.type === 'export') {
      await runAgentExport(action.export_type);
    }
  }
}

async function runAgentExport(exportType) {
  if (exportType === 'overview') {
    return downloadFile('/api/export/overview', { title: $('projectTitle').value.trim(), result: collectResultPayload(), follow_up: state.followUp, our_quotes: state.ourQuotes, competitor_quotes: state.competitorQuotes, timeline: state.timeline }, '解析总览.xlsx');
  }
  if (exportType === 'extraction') return downloadFile('/api/export/extraction', { result: collectResultPayload() }, '标书摘取.xlsx');
  if (exportType === 'register') return downloadFile('/api/export/register', { rows: currentRegisterRows(), sheet_name: $('sheetName').value }, '招标登记.xlsx');
  if (exportType === 'ledger') return downloadFile('/api/export/ledger', currentHistoryFilters(), '历史台账汇总.xlsx');
  if (exportType === 'our_quotes') return downloadFile('/api/export/our-quotes', { title: $('projectTitle').value.trim(), follow_up: state.followUp, rows: state.ourQuotes }, '我司报价一览表.xlsx');
  if (exportType === 'competitor_quotes') return downloadFile('/api/export/competitor-quotes', { title: $('projectTitle').value.trim(), follow_up: state.followUp, rows: state.competitorQuotes }, '竞对报价对比.xlsx');
  if (exportType === 'market_ledger') return downloadFile('/api/export/market-skill', { kind: state.marketKind, ...marketFilters() }, state.marketKind === 'cargo' ? '商机收集台账.xlsx' : '新造船收集台账.xlsx');
  if (exportType === 'market_report') {
    if (!state.marketReport) await generateMarketReport();
    return downloadFile('/api/export/market-skill-report', { kind: state.marketKind, report: state.marketReport, format: 'docx' }, state.marketKind === 'cargo' ? '商机市场AI分析报告.docx' : '新造船市场AI分析报告.docx');
  }
}

async function sendAgentMessage(forcedText = '') {
  const input = $('agentInput');
  const fileInput = $('agentFileInput');
  const text = (forcedText || input?.value || '').trim();
  const file = fileInput?.files?.[0] || null;
  if (!text && !file) return;
  const displayText = text || `上传文件：${file.name}`;
  state.agentTranscript.push({ role: 'user', content: displayText });
  renderAgentMessages();
  if (input) input.value = '';
  const formData = new FormData();
  formData.append('message', text);
  formData.append('conversation_id', state.agentConversationId || '');
  formData.append('page_context', JSON.stringify(agentPageContext()));
  if (file) formData.append('file', file);
  if (fileInput) fileInput.value = '';
  if ($('agentFileName')) $('agentFileName').textContent = '未选择附件';
  const button = $('agentSendBtn');
  if (button) {
    button.disabled = true;
    button.textContent = '处理中...';
  }
  try {
    const result = await request('/api/agent/chat', { method: 'POST', body: formData });
    state.agentConversationId = result.conversation_id || state.agentConversationId;
    localStorage.setItem('shipping_agent_conversation_id', state.agentConversationId);
    state.agentTranscript.push({ role: 'assistant', content: result.answer || '已处理。' });
    renderAgentMessages();
    setAgentPending(result.pending_change || {});
    await applyAgentActions(result.actions || []);
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = '发送';
    }
  }
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
            ${project.source_channel === 'feishu' ? '<span class="source-chip feishu">飞书录入</span>' : ''}
          </div>
          <small>${(project.updated_at || '').replace('T', ' ')}</small>
        </button>
      `
    )
    .join('');
  document.querySelectorAll('#projectList [data-project-id]').forEach((node) => {
    node.onclick = async () => {
      try {
        const projectId = node.dataset.projectId || '';
        if (!projectId) throw new Error('项目 ID 无效。');
        await loadProject(projectId);
      } catch (error) {
        showToast(error.message, true);
      }
    };
  });
}

function marketFields() {
  return state.marketKind === 'cargo' ? marketCargoFields : marketNewbuildingFields;
}

function emptyMarketRecord(kind = state.marketKind) {
  return kind === 'cargo'
    ? {
        board_type: '即时货盘',
        segment: '内贸化',
        cargo_name: '',
        tonnage: '',
        load_port: '',
        discharge_port: '',
        laycan: '',
        cargo_owner: '',
        cargo_date: new Date().toISOString().slice(0, 10),
        source: '手动录入',
        status: '跟进中',
        final_price: '',
        deal_date: '',
        deal_price: '',
        price_unit: '元/吨',
        currency: 'CNY',
        route: '',
        cargo_standard_name: '',
        executing_vessel: '',
        market_info: '',
        loss_reason: '',
        competitor_name: '',
        competitor_price: '',
        remark: '',
        raw_text: '',
      }
    : {
        stage: '信息待获取',
        ship_name: '',
        update_date: new Date().toISOString().slice(0, 10),
        shipyard: '',
        owner: '',
        dwt: '',
        build_status: '',
        delivery_time: '',
        actual_delivery_date: '',
        status_update_date: new Date().toISOString().slice(0, 10),
        status_note: '',
        contract_date: '',
        contract_price: '',
        ship_type: '',
        source: '手动录入',
        remark: '',
        raw_text: '',
      };
}

function setMarketRecord(record = null, { persisted = false } = {}) {
  state.marketRecord = { ...emptyMarketRecord(), ...(record || {}) };
  state.marketCurrentId = persisted ? record?.id || '' : '';
  if ($('marketRawText')) $('marketRawText').value = state.marketRecord.raw_text || '';
  if ($('marketReportPrompt') && !$('marketReportPrompt').value.trim()) $('marketReportPrompt').value = defaultMarketPrompt();
  renderMarketForm();
  setStatus('marketSaveStatus', state.marketCurrentId ? '已加载记录' : '新记录', state.marketCurrentId ? 'ok' : 'neutral');
  if ($('marketExtractBtn')) $('marketExtractBtn').textContent = state.marketCurrentId ? 'AI更新当前记录' : 'AI识别填表';
}

function marketInputMarkup(field, value) {
  if (field.type === 'select') {
    return `
      <select data-type="market-field" data-key="${field.key}">
        ${field.options.map((option) => `<option value="${option}" ${option === value ? 'selected' : ''}>${option}</option>`).join('')}
      </select>
    `;
  }
  if (field.type === 'textarea') {
    return `<textarea data-type="market-field" data-key="${field.key}" rows="3">${value || ''}</textarea>`;
  }
  return `<input data-type="market-field" data-key="${field.key}" value="${value || ''}" ${field.type === 'date' ? 'type="date"' : ''} />`;
}

function marketSourceLabel(item = {}) {
  return item.source_channel === 'feishu' || item.source === '飞书' ? '飞书录入' : item.source || '手动录入';
}

function marketSourceText(item = {}) {
  return item.source_text || item.raw_text || '';
}

function renderMarketSourceSummary() {
  const node = $('marketSourceSummary');
  if (!node) return;
  const record = state.marketRecord || {};
  const isFeishu = record.source_channel === 'feishu' || record.source === '飞书';
  const sourceText = marketSourceText(record);
  if (!isFeishu && !sourceText) {
    node.classList.add('hidden');
    node.innerHTML = '';
    return;
  }
  const meta = [
    marketSourceLabel(record),
    record.confirmed_at ? `确认：${record.confirmed_at.replace('T', ' ')}` : '',
    record.source_message_id ? `消息：${record.source_message_id}` : '',
  ].filter(Boolean);
  node.classList.remove('hidden');
  node.innerHTML = `
    <div class="market-source-head">
      ${meta.map((item) => `<span>${escapeHtml(item)}</span>`).join('')}
    </div>
    ${sourceText ? `<p>${escapeHtml(sourceText)}</p>` : ''}
  `;
}

function renderMarketForm() {
  if (!$('marketForm')) return;
  $('marketFormTitle').textContent = state.marketKind === 'cargo' ? '商机录入' : '新造船录入';
  $('marketLedgerTitle').textContent = state.marketKind === 'cargo' ? '商机台账' : '新造船台账';
  $('marketForm').innerHTML = marketFields()
    .map((field) => {
      const spanClass = field.type === 'textarea' ? ' field-span-2' : '';
      return `
        <label class="field${spanClass}">
          <span>${field.label}</span>
          ${marketInputMarkup(field, state.marketRecord[field.key])}
        </label>
      `;
    })
    .join('');
  document.querySelectorAll("[data-type='market-field']").forEach((node) => {
    const handler = (event) => {
      state.marketRecord[event.target.dataset.key] = event.target.value;
    };
    node.oninput = handler;
    node.onchange = handler;
  });
  renderMarketSourceSummary();
}

function renderMarketFilters() {
  document.querySelectorAll('.market-tab').forEach((node) => {
    node.classList.toggle('active', node.dataset.kind === state.marketKind);
  });
  document.querySelectorAll('.market-cargo-filter').forEach((node) => {
    node.classList.toggle('hidden', state.marketKind !== 'cargo');
  });
  document.querySelectorAll('.market-newbuilding-filter').forEach((node) => {
    node.classList.toggle('hidden', state.marketKind !== 'newbuilding');
  });
}

function renderMarketStats() {
  const stats = state.marketStats || {};
  const pairs =
    state.marketKind === 'cargo'
      ? [
          ['商机总数', stats.total],
          ['已成交', stats.won],
          ['未成交/放弃', stats.lost],
          ['跟进中', stats.tracking],
          ['已录成交价', stats.with_final_price],
        ]
      : [
          ['新船总数', stats.total],
          ['已投运', stats['已完造并出厂投运']],
          ['签约未开造', stats['合同签订未开造']],
          ['2027交付', stats['预计2027年完造出厂']],
          ['2028交付', stats['预计2028年完造出厂']],
          ['信息待获取', stats['信息待获取']],
        ];
  $('marketCount').textContent = `${stats.total || 0} 条`;
  $('marketStats').innerHTML = pairs
    .map(
      ([label, value]) => `
        <article class="history-stat">
          <span>${label}</span>
          <strong>${value || 0}</strong>
        </article>
      `
    )
    .join('');
}

function numberText(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number.toLocaleString('zh-CN') : value || '-';
}

function renderMarketSvg(points) {
  if (!points.length) {
    return `<div class="project-empty">暂无可绘制的成交价样本</div>`;
  }
  const values = points.map((point) => Number(point.avg_price || point.price)).filter((value) => Number.isFinite(value));
  if (!values.length) {
    return `<div class="project-empty">暂无可绘制的成交价样本</div>`;
  }
  const min = Math.min(...values);
  const max = Math.max(...values);
  const width = 760;
  const height = 240;
  const pad = 36;
  const span = max - min || 1;
  const coords = values.map((value, index) => {
    const x = pad + (index * (width - pad * 2)) / Math.max(values.length - 1, 1);
    const y = height - pad - ((value - min) * (height - pad * 2)) / span;
    return [x, y];
  });
  const polyline = coords.map(([x, y]) => `${x},${y}`).join(' ');
  return `
    <svg class="market-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="成交价走势图">
      <line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" />
      <line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" />
      <polyline points="${polyline}" />
      ${coords
        .map(([x, y], index) => `<circle cx="${x}" cy="${y}" r="4"><title>${escapeHtml(points[index].month || points[index].deal_date || '')} ${escapeHtml(numberText(values[index]))}</title></circle>`)
        .join('')}
      <text x="${pad}" y="22">${numberText(max)}</text>
      <text x="${pad}" y="${height - 8}">${numberText(min)}</text>
    </svg>
  `;
}

function renderMarketReport() {
  if (!$('marketAnalysis')) return;
  $('marketAnalysisTitle').textContent = state.marketKind === 'cargo' ? '商机市场分析' : '新造船市场分析';
  const report = state.marketReport;
  if (!report) {
    $('marketAnalysis').innerHTML = `<div class="project-empty">选择时间并点击生成报告后，这里会展示走势图、统计结论和 AI 建议。</div>`;
    return;
  }
  const stats = report.source_stats || {};
  const statPairs =
    state.marketKind === 'cargo'
      ? [
          ['纳入商机', stats.total],
          ['成交样本', stats.deal_count],
          ['平均成交价', stats.avg_price],
          ['成交率', `${stats.win_rate || 0}%`],
        ]
      : [
          ['新船信息', stats.total],
          ['已出厂/投运', stats.delivered],
          ['状态变化', stats.recent_change_count],
          ['交付年份', (stats.delivery_year_counts || []).length],
        ];
  const chartHtml =
    state.marketKind === 'cargo'
      ? renderMarketSvg(report.trend_points || [])
      : `<div class="market-bars">${(stats.stage_counts || [])
          .map((item) => `<div><span>${escapeHtml(item.name)}</span><strong style="width:${Math.max(8, item.count * 28)}px">${escapeHtml(item.count)}</strong></div>`)
          .join('') || '<div class="project-empty">暂无阶段统计</div>'}</div>`;
  const detailRows = report.detail_rows || [];
  const columns =
    state.marketKind === 'cargo'
      ? [
          ['deal_date', '成交日期'],
          ['cargo', '货品'],
          ['route', '航线'],
          ['price', '成交价'],
          ['price_unit', '单位'],
          ['currency', '币种'],
          ['executing_vessel', '最终执行船舶'],
          ['competitor_price', '竞对价'],
        ]
      : [
          ['status_update_date', '更新时间'],
          ['ship_name', '船名'],
          ['shipyard', '船厂'],
          ['owner', '船东'],
          ['stage', '阶段'],
          ['delivery_time', '预计交付'],
          ['actual_delivery_date', '实际出厂'],
        ];
  $('marketAnalysis').innerHTML = `
    <div class="history-stats market-report-stats">
      ${statPairs
        .map(
          ([label, value]) => `
            <article class="history-stat">
              <span>${escapeHtml(label)}</span>
              <strong>${escapeHtml(value ?? 0)}</strong>
            </article>
          `
        )
        .join('')}
    </div>
    <div class="market-report-layout">
      <div class="market-chart-card">${chartHtml}</div>
      <div class="analysis-card">
        <h4>AI / 规则结论</h4>
        <p>${escapeHtml(report.summary || '暂无结论')}</p>
        <h4>关键发现</h4>
        <ul>${(report.key_findings || []).map((item) => `<li>${escapeHtml(item)}</li>`).join('') || '<li>暂无</li>'}</ul>
        <h4>风险与建议</h4>
        <ul>${[...(report.risks || []), ...(report.recommendations || [])].map((item) => `<li>${escapeHtml(item)}</li>`).join('') || '<li>暂无</li>'}</ul>
      </div>
    </div>
    <div class="table-wrap compact-table">
      <table class="dense-table">
        <thead><tr>${columns.map(([, label]) => `<th>${escapeHtml(label)}</th>`).join('')}</tr></thead>
        <tbody>
          ${detailRows
            .slice(0, 20)
            .map((row) => `<tr>${columns.map(([key]) => `<td>${escapeHtml(row[key] ?? '-')}</td>`).join('')}</tr>`)
            .join('') || `<tr><td colspan="${columns.length}">暂无明细数据</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

function renderMarketList() {
  renderMarketFilters();
  renderMarketStats();
  renderMarketReport();
  if (!state.marketItems.length) {
    $('marketList').innerHTML = `<div class="project-empty">暂无匹配记录</div>`;
    return;
  }
  $('marketList').innerHTML = state.marketItems
    .map((item) => {
      const title =
        state.marketKind === 'cargo'
          ? `${item.cargo_name || '未填写货品'} ${item.load_port || ''} → ${item.discharge_port || ''}`
          : `${item.ship_name || item.owner || '未命名新船'} ${item.shipyard || ''}`;
      const chip = state.marketKind === 'cargo' ? item.status : item.stage;
      const sub =
        state.marketKind === 'cargo'
          ? `${item.tonnage || '-'} | ${item.cargo_owner || '-'} | ${item.final_price || '未录成交价'}${item.executing_vessel ? ` | 执行船：${item.executing_vessel}` : ''}`
          : `${item.dwt || '-'} | ${item.owner || '-'} | ${item.delivery_time || '-'}`;
      const sourceLabel = marketSourceLabel(item);
      const sourceText = marketSourceText(item);
      return `
        <button class="project-item market-item ${item.id === state.marketCurrentId ? 'active' : ''}" data-market-id="${item.id}">
          <div class="project-card-head">
            <strong>${title}</strong>
            <span class="status-chip ${chip === '已成交' || chip === '已完造并出厂投运' ? 'ok' : chip === '未成交' || chip === '放弃' ? 'error' : 'warn'}">${chip || '-'}</span>
          </div>
          <span>${sub}</span>
          <div class="project-tags">
            <span>${state.marketKind === 'cargo' ? item.board_type : item.build_status || '状态待补'}</span>
            <span>${state.marketKind === 'cargo' ? item.segment : item.ship_type || '船型待补'}</span>
            <span class="${item.source_channel === 'feishu' || item.source === '飞书' ? 'source-chip feishu' : 'source-chip'}">${escapeHtml(sourceLabel)}</span>
            <span>${(item.updated_at || '').replace('T', ' ')}</span>
          </div>
          ${sourceText ? `<small class="market-source-preview">${escapeHtml(sourceText)}</small>` : ''}
        </button>
      `;
    })
    .join('');
  document.querySelectorAll('[data-market-id]').forEach((node) => {
    node.onclick = async () => {
      try {
        const marketId = node.dataset.marketId || '';
        if (!marketId) throw new Error('市场情报记录 ID 无效。');
        const result = await request(`/api/market-skill/${state.marketKind}/${marketId}`);
        setMarketRecord(result.record, { persisted: true });
        renderMarketList();
      } catch (error) {
        showToast(error.message, true);
      }
    };
  });
}

function marketFilters() {
  return {
    q: $('marketSearch')?.value.trim() || '',
    board_type: state.marketKind === 'cargo' ? $('marketBoardType')?.value || '' : '',
    segment: state.marketKind === 'cargo' ? $('marketSegment')?.value || '' : '',
    status: state.marketKind === 'cargo' ? $('marketStatus')?.value || '' : '',
  };
}

function defaultMarketPrompt() {
  return state.marketKind === 'cargo'
    ? `请作为油化品航运市场分析师，基于当前筛选周期内的商机台账和成交价样本，形成一份可给经营层阅读的市场分析结论。
分析要求：
1. 先说明本次样本口径，包括纳入商机数量、已成交数量、有效成交价样本、主要货品和主要航线。
2. 只使用“已成交且有最终成交价”的数据判断市场运价走势，不要把意向价、我司报价或竞对报价混入主趋势。
3. 按货品、航线、月份分析成交均价变化，指出价格上行、下行或样本不足无法判断的原因。
4. 结合竞对价格和丢单原因，判断客户价格预期、竞争强度、船位匹配、TCE压力等经营风险。
5. 输出明确的经营建议，包括重点跟进航线/货品、报价策略、客户沟通重点、需要继续补充的数据字段。
6. 如果样本不足，请明确提示“样本不足”，不要过度推断。`
    : `请作为油化船新造船市场分析师，基于当前筛选周期内的新造船台账，形成一份可给经营层和投资决策参考的结论性报告。
分析要求：
1. 先说明本次样本口径，包括新造船记录数量、已出厂投运数量、签约未开造数量、预计交付年份分布。
2. 重点分析船舶建造状态变化、实际出厂投运节奏、未来集中交付年份和潜在供给释放压力。
3. 按船厂、船东、船型、DWT和交付时间识别值得关注的新增供给或延期风险。
4. 从供给侧变化推导对油化品船运价、船舶投资、租船经营和市场竞争的可能影响。
5. 输出明确建议，包括需要重点跟踪的船厂/船东、投资节奏建议、市场风险预警、下一阶段需补充的信息。
6. 如果样本不足，请明确提示“样本不足”，不要编造外部订单或价格信息。`;
}

function marketReportPayload() {
  return {
    kind: state.marketKind,
    period: 'custom',
    start_date: $('marketReportStart')?.value || '',
    end_date: $('marketReportEnd')?.value || '',
    custom_prompt: $('marketReportPrompt')?.value.trim() || defaultMarketPrompt(),
    filters: marketFilters(),
  };
}

async function loadMarketRecords() {
  const requestSeq = ++state.marketRequestSeq;
  const params = new URLSearchParams(marketFilters());
  const result = await request(`/api/market-skill/${state.marketKind}?${params.toString()}`);
  if (requestSeq !== state.marketRequestSeq) return;
  state.marketItems = result.items || [];
  state.marketStats = result.stats || {};
  renderMarketList();
}

let marketSearchTimer = null;

function scheduleMarketRecordsLoad() {
  clearTimeout(marketSearchTimer);
  marketSearchTimer = setTimeout(async () => {
    try {
      await loadMarketRecords();
    } catch (error) {
      showToast(error.message, true);
    }
  }, 180);
}

async function extractMarketRecord() {
  const text = $('marketRawText').value.trim();
  if (!text) throw new Error('请先粘贴要识别的市场情报文本');
  const result = await request('/api/market-skill/extract', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      kind: state.marketKind,
      text,
      record_id: state.marketCurrentId,
      current_record: { ...state.marketRecord, raw_text: text },
    }),
  });
  setMarketRecord(result.record, { persisted: result.mode === 'update' });
  showToast(result.mode === 'update' ? 'AI 已更新当前记录表单，请确认后保存' : '已识别并填入表单，请人工确认后保存');
}

async function saveMarketRecord() {
  const rawText = $('marketRawText').value.trim();
  if (state.marketKind === 'cargo') {
    state.marketRecord.route = state.marketRecord.route || [state.marketRecord.load_port, state.marketRecord.discharge_port].filter(Boolean).join(' - ');
    state.marketRecord.cargo_standard_name = state.marketRecord.cargo_standard_name || state.marketRecord.cargo_name || '';
    state.marketRecord.deal_price = state.marketRecord.deal_price || state.marketRecord.final_price || '';
  }
  const payload = { kind: state.marketKind, record: { ...state.marketRecord, raw_text: rawText } };
  const url = state.marketCurrentId ? `/api/market-skill/${state.marketKind}/${state.marketCurrentId}` : `/api/market-skill/${state.marketKind}`;
  const result = await request(url, {
    method: state.marketCurrentId ? 'PUT' : 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  setMarketRecord(result.record, { persisted: true });
  state.marketReport = null;
  await loadMarketRecords();
  showToast('市场情报记录已保存');
}

async function deleteMarketRecord() {
  if (!state.marketCurrentId) throw new Error('当前没有已保存的市场情报记录');
  if (!window.confirm('确认删除当前市场情报记录吗？删除后无法恢复。')) return;
  await request(`/api/market-skill/${state.marketKind}/${state.marketCurrentId}`, { method: 'DELETE' });
  setMarketRecord();
  await loadMarketRecords();
  showToast('市场情报记录已删除');
}

async function generateMarketReport() {
  const button = $('marketGenerateReportBtn');
  const oldText = button?.textContent || '生成报告';
  try {
    if (button) {
      button.disabled = true;
      button.textContent = 'AI分析中...';
    }
    showToast('AI 正在汇总台账、计算趋势并生成分析报告');
    const report = await request('/api/market-skill/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(marketReportPayload()),
    });
    state.marketReport = report;
    renderMarketReport();
    document.querySelector('.market-analysis-card')?.setAttribute('open', '');
    showToast('市场分析报告已生成，可以预览或导出 Word');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = oldText;
    }
  }
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
  if ($('feishuCallbackUrl')) {
    $('feishuCallbackUrl').value = `${window.location.origin}/api/feishu/events`;
    $('feishuEnabled').checked = Boolean(config.feishu_enabled);
    $('feishuReceiveMode').value = config.feishu_receive_mode || 'ws';
    $('feishuAppId').value = config.feishu_app_id || '';
    $('feishuAppSecret').value = '';
    $('feishuVerificationToken').value = '';
    $('feishuEncryptKey').value = '';
    $('feishuAllowedOpenIds').value = config.feishu_allowed_open_ids || '';
    $('feishuAllowedChatIds').value = config.feishu_allowed_chat_ids || '';
    setStatus('feishuStatus', config.feishu_enabled ? '已启用' : '未启用', config.feishu_enabled ? 'ok' : 'neutral');
  }
  setStatus('configStatus', config.base_url && config.has_api_key ? '已保存' : '未保存', config.base_url && config.has_api_key ? 'ok' : 'neutral');
}

function readConfigForm() {
  return {
    base_url: $('baseUrl').value.trim(),
    api_key: $('apiKey').value.trim(),
    model: $('modelName').value.trim(),
    temperature: Number($('temperature').value || 0.1),
    feishu_enabled: Boolean($('feishuEnabled')?.checked),
    feishu_receive_mode: $('feishuReceiveMode')?.value || 'ws',
    feishu_app_id: $('feishuAppId')?.value.trim() || '',
    feishu_app_secret: $('feishuAppSecret')?.value.trim() || '',
    feishu_verification_token: $('feishuVerificationToken')?.value.trim() || '',
    feishu_encrypt_key: $('feishuEncryptKey')?.value.trim() || '',
    feishu_allowed_open_ids: $('feishuAllowedOpenIds')?.value.trim() || '',
    feishu_allowed_chat_ids: $('feishuAllowedChatIds')?.value.trim() || '',
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
  const config = await request('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(readConfigForm()),
  });
  setStatus('configStatus', '已保存', 'ok');
  if ($('feishuStatus')) {
    setStatus('feishuStatus', config.feishu_enabled ? '已启用' : '未启用', config.feishu_enabled ? 'ok' : 'neutral');
  }
  showToast('配置已保存');
}

function hasConfigFormValues() {
  return Boolean($('baseUrl').value.trim() && $('modelName').value.trim() && $('apiKey').value.trim());
}

async function testConfig() {
  const result = await request('/api/config/test', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(readConfigForm()),
  });
  showToast(`连通成功：${result.message}`);
}

async function testFeishuMessage() {
  await saveConfig();
  await request('/api/feishu/test-message', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      receive_id_type: $('feishuTestReceiveType').value,
      receive_id: $('feishuTestReceiveId').value.trim(),
      text: '标书解析机器人测试消息：配置已连通。',
    }),
  });
  showToast('飞书测试消息已发送');
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
  if (hasConfigFormValues()) {
    await saveConfig();
  }
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
  state.projectSource = {
    source_channel: project.source_channel || '',
    source_open_id: project.source_open_id || '',
    source_chat_id: project.source_chat_id || '',
    source_message_id: project.source_message_id || '',
    source_text: project.source_text || '',
    confirmed_at: project.confirmed_at || '',
  };
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
  document.querySelectorAll('.main-tab').forEach((node) => {
    node.onclick = () => {
      switchWorkspace(node.dataset.workspace);
      if (node.dataset.workspace === 'bidWorkspace') {
        const activePanel = document.querySelector('#bidWorkspace .nav-link.active')?.dataset.target || 'overviewPanel';
        switchPanel(activePanel);
      }
    };
  });

  document.querySelectorAll('.nav-link').forEach((node) => {
    node.onclick = () => {
      if (node.dataset.target === 'chatPanel') {
        toggleAgentDrawer(true);
        return;
      }
      switchPanel(node.dataset.target);
    };
  });

  bindClick('settingsToggleBtn', () => toggleSettings());
  bindClick('settingsCloseBtn', () => toggleSettings(false));
  bindClick('agentToggleBtn', () => toggleAgentDrawer());
  bindClick('agentCloseBtn', () => toggleAgentDrawer(false));
  bindClick('agentSendBtn', async () => {
    try {
      await sendAgentMessage();
    } catch (error) {
      state.agentTranscript.push({ role: 'assistant', content: `处理失败：${error.message}` });
      renderAgentMessages();
      showToast(error.message, true);
    }
  });
  bindChange('agentFileInput', () => {
    const file = $('agentFileInput')?.files?.[0];
    if ($('agentFileName')) $('agentFileName').textContent = file?.name || '未选择附件';
  });
  $('agentInput')?.addEventListener('keydown', async (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
      event.preventDefault();
      $('agentSendBtn')?.click();
    }
  });

  document.querySelectorAll('.market-analysis-summary button').forEach((node) => {
    node.addEventListener('click', (event) => event.stopPropagation());
  });

  ['projectSearch', 'historyYear', 'historyBidStatus', 'historyAwardStatus'].forEach((id) => {
    const handler = async () => {
      try {
        await loadProjects();
      } catch (error) {
        showToast(error.message, true);
      }
    };
    bindInput(id, handler, true);
  });

  ['marketSearch', 'marketBoardType', 'marketSegment', 'marketStatus'].forEach((id) => {
    bindInput(
      id,
      () => {
        state.marketReport = null;
        renderMarketReport();
        scheduleMarketRecordsLoad();
      },
      true
    );
  });

  ['marketReportStart', 'marketReportEnd', 'marketReportPrompt'].forEach((id) => {
    bindInput(
      id,
      () => {
        state.marketReport = null;
        renderMarketReport();
      },
      true
    );
  });

  document.querySelectorAll('.market-tab').forEach((node) => {
    node.onclick = async () => {
      state.marketKind = node.dataset.kind;
      state.marketCurrentId = '';
      state.marketReport = null;
      if ($('marketReportPrompt')) $('marketReportPrompt').value = defaultMarketPrompt();
      setMarketRecord();
      try {
        await loadMarketRecords();
      } catch (error) {
        showToast(error.message, true);
      }
    };
  });

  bindChange('configPreset', (event) => applyPreset(event.target.value));

  bindClick('logoutBtn', async () => {
    await logout();
  });

  bindClick('saveConfigBtn', async () => {
    try {
      await saveConfig();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('testConfigBtn', async () => {
    try {
      await testConfig();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('saveFeishuBtn', async () => {
    try {
      await saveConfig();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('testFeishuBtn', async () => {
    try {
      await testFeishuMessage();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('parseBtn', async () => {
    const button = $('parseBtn');
    if (!button) return;
    try {
      button.disabled = true;
      button.textContent = '解析中...';
      await parseFile();
    } catch (error) {
      setStatus('fileStatus', '解析失败', 'error');
      showToast(error.message, true);
    } finally {
      button.disabled = false;
      button.textContent = '开始解析';
    }
  });

  bindClick('saveProjectBtn', async () => {
    try {
      await saveCurrentProject();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('deleteProjectBtn', async () => {
    try {
      await deleteCurrentProject();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindChange('registerMode', () => {
    const registerMode = $('registerMode');
    if (!registerMode) return;
    state.registerMode = registerMode.value;
    if (!$('resultSection').classList.contains('hidden')) {
      renderRegister();
      bindDynamicEditors();
    }
  });

  bindInput('sheetName', () => {
    const sheetName = $('sheetName');
    if (!sheetName) return;
    state.followUp.register_year = sheetName.value.trim();
  });

  bindClick('addOurQuoteBtn', () => {
    state.ourQuotes.push(newQuote('我司'));
    renderTracking();
    bindDynamicEditors();
    renderSummary(state.documentSummary);
  });

  bindClick('addCompetitorQuoteBtn', () => {
    state.competitorQuotes.push(newQuote(''));
    renderTracking();
    bindDynamicEditors();
    renderSummary(state.documentSummary);
  });

  bindClick('addTimelineBtn', () => {
    state.timeline.push(newTimeline());
    renderTracking();
    bindDynamicEditors();
  });

  bindClick('sendChatBtn', async () => {
    const chatInput = $('chatInput');
    if (!chatInput) return;
    const question = chatInput.value.trim();
    if (!question) return;
    chatInput.value = '';
    try {
      await sendChat(question);
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('askDefaultBtn', async () => {
    try {
      toggleAgentDrawer(true);
      await sendAgentMessage('请重新总结当前标书的关键节点、资格要求、报价要求、评标标准和主要风险。');
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('exportOverviewBtn', async () => {
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
  });

  bindClick('exportExtractionBtn', async () => {
    try {
      await downloadFile('/api/export/extraction', { result: collectResultPayload() }, '标书摘取.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('exportRegisterBtn', async () => {
    try {
      await downloadFile('/api/export/register', { rows: currentRegisterRows(), sheet_name: $('sheetName').value }, '招标登记.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('exportOurQuotesBtn', async () => {
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
  });

  bindClick('exportCompetitorQuotesBtn', async () => {
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
  });

  bindClick('exportLedgerBtn', async () => {
    try {
      await downloadFile('/api/export/ledger', currentHistoryFilters(), '历史台账汇总.xlsx');
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketExtractBtn', async () => {
    try {
      await extractMarketRecord();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketNewBtn', () => {
    setMarketRecord();
  });

  bindClick('marketSaveBtn', async () => {
    try {
      await saveMarketRecord();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketDeleteBtn', async () => {
    try {
      await deleteMarketRecord();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketExportBtn', async () => {
    try {
      await downloadFile(
        '/api/export/market-skill',
        { kind: state.marketKind, ...marketFilters() },
        state.marketKind === 'cargo' ? '商机收集台账.xlsx' : '新造船收集台账.xlsx'
      );
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketGenerateReportBtn', async () => {
    try {
      await generateMarketReport();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('marketReportExportBtn', async () => {
    try {
      if (!state.marketReport) {
        await generateMarketReport();
      }
      await downloadFile(
        '/api/export/market-skill-report',
        { kind: state.marketKind, report: state.marketReport, format: 'docx' },
        state.marketKind === 'cargo' ? '商机市场AI分析报告.docx' : '新造船市场AI分析报告.docx'
      );
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('createUserBtn', async () => {
    try {
      await createUser();
    } catch (error) {
      showToast(error.message, true);
    }
  });

  bindClick('refreshUsersBtn', async () => {
    try {
      await loadUsers();
    } catch (error) {
      showToast(error.message, true);
    }
  });
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
    setMarketRecord();
    await loadProjects();
    await loadMarketRecords();
    await loadAgentSession();
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
