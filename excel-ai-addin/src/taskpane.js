/* ── 审计AI助手 — 核心逻辑 ── */

// ── 状态 ──
const state = {
  messages: [],
  dataContext: null,
  dataRange: null,
  isProcessing: false,
  model: 'deepseek-v4-pro',
  serverUrl: 'http://localhost:3000',
  systemPrompt: '你是资深审计师，精通数据分析、异常检测和对账处理。分析数据时，直接给出结论和具体数字，不空泛。数据以表格形式呈现时，逐行解释异常值。回答简洁、专业、可操作。',
  lastAssistantContent: '',
};

// ── DOM 引用 ──
const $ = (id) => document.getElementById(id);
const dom = {
  messages: $('messages'),
  welcome: $('welcome'),
  inputBox: $('inputBox'),
  btnSend: $('btnSend'),
  btnReadSelection: $('btnReadSelection'),
  btnClearData: $('btnClearData'),
  btnWrite: $('btnWrite'),
  btnClear: $('btnClear'),
  btnSettings: $('btnSettings'),
  btnCloseSettings: $('btnCloseSettings'),
  btnSaveSettings: $('btnSaveSettings'),
  btnRetry: $('btnRetry'),
  statusDot: $('statusDot'),
  dataInfo: $('dataInfo'),
  modelBadge: $('modelBadge'),
  connWarning: $('connWarning'),
  settingsPanel: $('settingsPanel'),
  modelSelect: $('modelSelect'),
  serverUrl: $('serverUrl'),
  systemPrompt: $('systemPrompt'),
  chatArea: $('chatArea'),
};

// ── 工具函数 ──

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function renderMarkdown(text) {
  if (!text) return '';
  let html = escapeHtml(text);

  // 代码块 (```)
  html = html.replace(/```(\w*)\n([\s\S]*?)```/g, (_, lang, code) => {
    return `<pre><code>${code}</code></pre>`;
  });

  // 行内代码 (`code`)
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');

  // 表格: | a | b |\n |---|---|\n | 1 | 2 |
  html = html.replace(/\n(\|.+\|\n\|[-| ]+\|\n(?:\|.+\|\n?)*)/g, (_, table) => {
    const rows = table.trim().split('\n');
    const headers = rows[0].split('|').filter(c => c.trim());
    let tbl = '<table><thead><tr>';
    headers.forEach(h => { tbl += `<th>${h.trim()}</th>`; });
    tbl += '</tr></thead><tbody>';
    for (let i = 2; i < rows.length; i++) {
      const cells = rows[i].split('|').filter(c => c.trim());
      if (cells.length === 0) continue;
      tbl += '<tr>';
      cells.forEach(c => { tbl += `<td>${c.trim()}</td>`; });
      tbl += '</tr>';
    }
    tbl += '</tbody></table>';
    return '\n' + tbl;
  });

  // 粗体 **text**
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

  // 列表: - item
  html = html.replace(/\n- (.+)/g, '\n<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');

  // 有序列表: 1. item
  html = html.replace(/\n(\d+)\. (.+)/g, '\n<li value="$1">$2</li>');

  // 换行
  html = html.replace(/\n{2,}/g, '</p><p>');
  html = html.replace(/\n/g, '<br>');

  return `<p>${html}</p>`;
}

// ── 消息管理 ──

function addMessage(role, content, extraClass = '') {
  const div = document.createElement('div');
  div.className = `message ${role} ${extraClass}`;

  if (role === 'system') {
    // 数据上下文用 pre 展示
    const pre = document.createElement('pre');
    pre.textContent = content;
    div.appendChild(pre);
  } else if (role === 'assistant') {
    div.innerHTML = renderMarkdown(content);
  } else {
    div.textContent = content;
  }

  dom.welcome.classList.add('hidden');
  dom.messages.appendChild(div);
  scrollToBottom();
  return div;
}

function addTypingIndicator() {
  const div = document.createElement('div');
  div.className = 'typing-indicator';
  div.id = 'typingIndicator';
  for (let i = 0; i < 3; i++) {
    const dot = document.createElement('span');
    dot.className = 'typing-dot';
    div.appendChild(dot);
  }
  dom.messages.appendChild(div);
  scrollToBottom();
  return div;
}

function removeTypingIndicator() {
  const el = document.getElementById('typingIndicator');
  if (el) el.remove();
}

function scrollToBottom() {
  requestAnimationFrame(() => {
    dom.chatArea.scrollTop = dom.chatArea.scrollHeight;
  });
}

// ── API 通信 ──

async function sendMessage(userText) {
  if (state.isProcessing) return;
  if (!userText.trim()) return;

  state.isProcessing = true;
  dom.btnSend.disabled = true;

  // 添加用户消息
  addMessage('user', userText);
  state.messages.push({ role: 'user', content: userText });

  // 打字指示器
  addTypingIndicator();

  try {
    const payload = {
      model: state.model,
      messages: state.messages,
      system: state.systemPrompt,
      max_tokens: 8192,
    };

    const resp = await fetch(`${state.serverUrl}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    removeTypingIndicator();

    if (!resp.ok) {
      const errData = await resp.json().catch(() => ({}));
      const errMsg = errData.error || `请求失败 (HTTP ${resp.status})`;
      addMessage('error', `AI 服务错误: ${errMsg}`);
      state.messages.push({ role: 'assistant', content: `[错误] ${errMsg}` });
      return;
    }

    const data = await resp.json();

    // 解析 Anthropic 格式响应
    let replyText = '';
    if (data.content && Array.isArray(data.content)) {
      replyText = data.content
        .filter(b => b.type === 'text')
        .map(b => b.text)
        .join('\n');
    } else if (data.content && typeof data.content === 'string') {
      replyText = data.content;
    } else {
      replyText = JSON.stringify(data);
    }

    if (!replyText) replyText = '(空响应)';

    addMessage('assistant', replyText);
    state.messages.push({ role: 'assistant', content: replyText });
    state.lastAssistantContent = replyText;

  } catch (err) {
    removeTypingIndicator();
    addMessage('error', `网络错误: ${err.message}。请检查本地服务是否运行。`);
  } finally {
    state.isProcessing = false;
    dom.btnSend.disabled = false;
    dom.inputBox.focus();
  }
}

// ── Excel 数据操作 ──

async function readSelection() {
  if (typeof Office === 'undefined' || !Office.context) {
    addMessage('system', 'Office.js 未加载，无法读取 Excel 数据');
    return;
  }

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('address, values, columnCount, rowCount');
      await context.sync();

      if (!range.values || range.values.length === 0) {
        addMessage('system', '选区为空，请选择包含数据的单元格');
        return;
      }

      const vals = range.values;
      const rows = vals.length;
      const cols = vals[0].length;

      // 格式化数据上下文
      let mdTable = '';
      // 第一行作为表头
      const headers = vals[0].map((_, i) => `列${i + 1}`);
      // 尝试使用第一行作为表头（如果包含文本）
      const firstRow = vals[0];
      const hasHeader = firstRow.some(v => typeof v === 'string' && isNaN(Number(v)));
      const startRow = hasHeader ? 1 : 0;

      if (hasHeader) {
        mdTable = '| ' + firstRow.map(h => String(h == null ? '' : h)).join(' | ') + ' |\n';
        mdTable += '| ' + firstRow.map(() => '---').join(' | ') + ' |\n';
      } else {
        mdTable = '| ' + headers.join(' | ') + ' |\n';
        mdTable += '| ' + headers.map(() => '---').join(' | ') + ' |\n';
      }

      for (let i = startRow; i < rows; i++) {
        mdTable += '| ' + vals[i].map(v => String(v == null ? '' : v)).join(' | ') + ' |\n';
      }

      const summary = `选区: ${range.address} (${rows}行 x ${cols}列)`;
      const contextMsg = `当前选中数据 (${range.address}):\n\`\`\`\n${mdTable}\n\`\`\``;

      state.dataContext = contextMsg;
      state.dataRange = range.address;

      dom.dataInfo.textContent = `${range.address} — ${rows}行 ${cols}列`;
      addMessage('system', contextMsg);
    });
  } catch (err) {
    addMessage('error', `读取数据失败: ${err.message}`);
  }
}

async function writeToSheet(content) {
  if (typeof Office === 'undefined' || !Office.context) {
    addMessage('error', 'Office.js 未加载，无法写入表格');
    return;
  }

  const textToWrite = content || state.lastAssistantContent;
  if (!textToWrite) {
    addMessage('system', '没有可写入的内容。先向 AI 提问获取回复。');
    return;
  }

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('address');
      await context.sync();

      // 如果选中了单个单元格，写文本
      // 如果选中了多行多列，按行列拆分
      const lines = textToWrite.trim().split('\n');
      const values = lines.map(line => [line]);

      range.values = values;
      range.format.autofitColumns();
      await context.sync();

      addMessage('system', `已写入 ${range.address} (${values.length} 行)`);
    });
  } catch (err) {
    addMessage('error', `写入失败: ${err.message}`);
  }
}

// ── 连接检查 ──

async function checkConnection() {
  try {
    const resp = await fetch(`${state.serverUrl}/api/check`, {
      method: 'POST',
      signal: AbortSignal.timeout(3000),
    });
    const data = await resp.json();
    if (data.proxy) {
      dom.statusDot.className = 'status-dot connected';
      dom.connWarning.classList.add('hidden');
    } else {
      dom.statusDot.className = 'status-dot error';
      dom.connWarning.classList.remove('hidden');
    }
  } catch {
    dom.statusDot.className = 'status-dot error';
    dom.connWarning.classList.remove('hidden');
  }
}

// ── 设置管理 ──

function loadSettings() {
  try {
    const saved = localStorage.getItem('audit_ai_settings');
    if (saved) {
      const s = JSON.parse(saved);
      if (s.model) state.model = s.model;
      if (s.serverUrl) state.serverUrl = s.serverUrl;
      if (s.systemPrompt) state.systemPrompt = s.systemPrompt;
    }
  } catch {}
  dom.modelSelect.value = state.model;
  dom.serverUrl.value = state.serverUrl;
  dom.systemPrompt.value = state.systemPrompt;
  dom.modelBadge.textContent = state.model;
}

function saveSettings() {
  state.model = dom.modelSelect.value;
  state.serverUrl = dom.serverUrl.value.replace(/\/+$/, '');
  state.systemPrompt = dom.systemPrompt.value;
  dom.modelBadge.textContent = state.model;
  try {
    localStorage.setItem('audit_ai_settings', JSON.stringify({
      model: state.model,
      serverUrl: state.serverUrl,
      systemPrompt: state.systemPrompt,
    }));
  } catch {}
  dom.settingsPanel.classList.add('hidden');
  checkConnection();
}

// ── 清空对话 ──

function clearChat() {
  state.messages = [];
  state.lastAssistantContent = '';
  dom.messages.innerHTML = '';
  dom.welcome.classList.remove('hidden');
}

function clearDataContext() {
  state.dataContext = null;
  state.dataRange = null;
  dom.dataInfo.textContent = '未读取';
}

// ── 初始化 ──

function init() {
  // Office.js
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      console.log('Office.js 已就绪 (Excel)');
    } else if (info.host === Office.HostType.Workbook) {
      console.log('Office.js 已就绪 (WPS)');
    } else {
      console.log('Office.js 已就绪, host:', info.host);
    }
  });

  loadSettings();
  checkConnection();

  // 每 30 秒检查连接
  setInterval(checkConnection, 30000);

  // ── 事件绑定 ──

  // 发送
  dom.btnSend.addEventListener('click', () => {
    const text = dom.inputBox.value.trim();
    if (text) {
      sendMessage(text);
      dom.inputBox.value = '';
    }
  });

  // 回车发送
  dom.inputBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const text = dom.inputBox.value.trim();
      if (text && !state.isProcessing) {
        sendMessage(text);
        dom.inputBox.value = '';
      }
    }
  });

  // 输入时启用/禁用发送按钮
  dom.inputBox.addEventListener('input', () => {
    dom.btnSend.disabled = !dom.inputBox.value.trim() || state.isProcessing;
  });

  // 读取选区
  dom.btnReadSelection.addEventListener('click', readSelection);

  // 清除数据上下文
  dom.btnClearData.addEventListener('click', clearDataContext);

  // 写入表格
  dom.btnWrite.addEventListener('click', () => writeToSheet());

  // 清空对话
  dom.btnClear.addEventListener('click', clearChat);

  // 设置面板
  dom.btnSettings.addEventListener('click', () => {
    dom.settingsPanel.classList.remove('hidden');
  });
  dom.btnCloseSettings.addEventListener('click', () => {
    dom.settingsPanel.classList.add('hidden');
  });
  dom.btnSaveSettings.addEventListener('click', saveSettings);

  // 重试连接
  dom.btnRetry.addEventListener('click', checkConnection);

  // 点击空白区域关闭设置
  dom.settingsPanel.addEventListener('click', (e) => {
    if (e.target === dom.settingsPanel) {
      dom.settingsPanel.classList.add('hidden');
    }
  });
}

// ── 启动 ──
document.addEventListener('DOMContentLoaded', init);
