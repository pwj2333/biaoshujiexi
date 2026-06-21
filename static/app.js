const state = {
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
  registerMode: 'packages',
  templateMeta: null,
  projects: [],
};

const presets = {
  custom: { base_url: '', model: '' },
  openai: { base_url: 'https://api.openai.com/v1', model: 'gpt-4.1-mini' },
  openrouter: { base_url: 'https://openrouter.ai/api/v1', model: 'openai/gpt-4.1-mini' },
};

const $ = (id) => document.getElementById(id);

async function request(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    let message = '请求失败';
    try {
      const data = await response.json();
      message = data.detail || message;
    } catch (_) {}
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
  return {
    title: $('projectTitle').value.trim(),
    source_file_name: state.sourceFileName,
    register_mode: $('registerMode').value,
    sheet_name: $('sheetName').value.trim(),
    result: collectResultPayload(),
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
  state.registerMode = $('registerMode').value;
  $('projectTitle').value = '';
  $('fileInput').value = '';
  $('chatMessages').innerHTML = '';
  $('summaryPanel').innerHTML = '';
  $('analysisPanel').innerHTML = '';
  $('reviewPanel').innerHTML = '';
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
                <td><textarea data-type="extraction-source" data-key="${field.key}" rows="5">${value.source_excerpt || ''}</textarea></td>
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
  renderExtraction();
  renderRegister();
  bindDynamicEditors();
  $('resultEmpty').classList.add('hidden');
  $('resultSection').classList.remove('hidden');
}

function renderProjectList() {
  $('projectCount').textContent = `${state.projects.length} 条`;
  if (!state.projects.length) {
    $('projectList').innerHTML = `<div class="project-empty">没有匹配的历史记录</div>`;
    return;
  }
  $('projectList').innerHTML = state.projects
    .map(
      (project) => `
        <button class="project-item ${project.project_id === state.currentProjectId ? 'active' : ''}" data-project-id="${project.project_id}">
          <strong>${project.title || '未命名项目'}</strong>
          <span>${project.bid_no || project.project_name || project.source_file_name || '暂无编号'}</span>
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
  const params = new URLSearchParams({ q: $('projectSearch').value.trim() });
  const result = await request(`/api/projects?${params.toString()}`);
  state.projects = result.items || [];
  renderProjectList();
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
  $('registerMode').value = state.registerMode;
  $('sheetName').value = project.sheet_name || $('sheetName').value;
  $('projectTitle').value = project.title || '';
  $('chatMessages').innerHTML = '';
  applyResult(project.result);
  $('historyMenu').open = true;
  setStatus('fileStatus', '已加载历史记录', 'ok');
  renderProjectList();
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

function bindEvents() {
  document.querySelectorAll('.nav-link').forEach((node) => {
    node.onclick = () => switchPanel(node.dataset.target);
  });

  $('historyMenu').addEventListener('toggle', () => {
    const toggle = document.querySelector('.history-toggle');
    if (toggle) {
      toggle.textContent = $('historyMenu').open ? '收起' : '展开';
    }
  });

  $('projectSearch').oninput = async () => {
    try {
      $('historyMenu').open = true;
      await loadProjects();
    } catch (error) {
      showToast(error.message, true);
    }
  };

  $('configPreset').onchange = (event) => applyPreset(event.target.value);

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
      await downloadFile('/api/export/overview', { result: collectResultPayload() }, '解析总览.md');
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
}

async function init() {
  try {
    await loadMeta();
    await loadConfig();
    bindEvents();
    clearCurrentProject();
    await loadProjects();
    switchPanel('overviewPanel');
  } catch (error) {
    showToast(error.message, true);
  }
}

init();
