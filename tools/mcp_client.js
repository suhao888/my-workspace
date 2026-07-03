#!/usr/bin/env node
/**
 * MCP 客户端 — 通过 CLI 调用三个 MCP 服务器（mineru/playwright/vision）
 * 用法:
 *   node mcp_client.js <server> <tool> [args...]
 *   node mcp_client.js <server> --list
 *
 * 示例:
 *   node mcp_client.js mineru --list
 *   node mcp_client.js playwright --list
 *   node mcp_client.js vision --list
 *
 *   node mcp_client.js mineru parse_pdf --file="C:/path/to/doc.pdf"
 *   node mcp_client.js playwright navigate --url="https://example.com"
 *   node mcp_client.js vision analyze_image --image="C:/path/to/img.png"
 */

const { spawn } = require('child_process');

// === 服务器配置 ===
const MCP_SERVERS = {
  mineru: {
    command: 'mineru-mcp',
    args: [],
    shell: true  // 无启动参数，shell:true 没问题
  },
  playwright: {
    command: 'playwright-mcp',
    // 用单个字符串构造命令行，传给 exec 或 shell spawn
    cmdline: 'playwright-mcp --browser msedge --executable-path "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" --headless'
  },
  vision: {
    command: 'vision-mcp-server',
    args: [],
    shell: true,
    env: {
      VISION_MCP_APPS_ROOT: 'E:\\Projects\\my-workspace\\apps'
    }
  }
};

// === JSON-RPC 工具 ===
let msgId = 1;
function createRequest(method, params = {}) {
  return JSON.stringify({
    jsonrpc: '2.0',
    id: msgId++,
    method,
    params
  }) + '\n';
}

// === 主逻辑 ===
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('用法: node mcp_client.js <server> <--list | tool> [args...]');
    console.error('服务器: mineru, playwright, vision');
    process.exit(1);
  }

  const serverName = args[0];
  const action = args[1];
  const toolArgs = args.slice(2);

  const config = MCP_SERVERS[serverName];
  if (!config) {
    console.error(`未知服务器: ${serverName}。可用: ${Object.keys(MCP_SERVERS).join(', ')}`);
    process.exit(1);
  }

  // 启动服务器进程
  console.error(`[MCP] 正在连接 ${serverName}...`);
  let proc;
  if (config.cmdline) {
    // 有完整命令行的（如 playwright-mcp）：将整个命令传给 shell
    proc = spawn(config.cmdline, [], { stdio: ['pipe', 'pipe', 'pipe'], shell: true, env: { ...process.env, ...(config.env || {}) } });
  } else {
    proc = spawn(config.command, config.args || [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      shell: config.shell !== false,
      env: { ...process.env, ...(config.env || {}) }
    });
  }

  let buffer = '';
  let initialized = false;
  let pendingResolve = null;
  let pendingReject = null;
  let responseTimer = null;

  // 解析服务器响应
  function processResponse(data) {
    buffer += data.toString();
    const lines = buffer.split('\n');
    buffer = lines.pop() || '';

    for (const line of lines) {
      if (!line.trim()) continue;
      try {
        const msg = JSON.parse(line);
        handleMessage(msg);
      } catch (e) {
        // 不完整消息，继续等待
      }
    }
  }

  function handleMessage(msg) {
    if (msg.id && pendingResolve) {
      const resolve = pendingResolve;
      const timer = responseTimer;
      pendingResolve = null;
      pendingReject = null;
      responseTimer = null;
      if (timer) clearTimeout(timer);
      resolve(msg);
    }
  }

  proc.stdout.on('data', processResponse);
  proc.stderr.on('data', (data) => {
    // 将 stderr 输出到调试
    process.stderr.write(`[${serverName}:stderr] ${data}`);
  });

  // 等待响应的辅助函数
  function waitForResponse(timeout = 30000) {
    return new Promise((resolve, reject) => {
      pendingResolve = resolve;
      pendingReject = reject;
      responseTimer = setTimeout(() => {
        if (pendingResolve) {
          const r = pendingResolve;
          pendingResolve = null;
          pendingReject = null;
          r({ error: { message: '请求超时' } });
        }
      }, timeout);
    });
  }

  try {
    // === 步骤 1: Initialize ===
    const initReq = createRequest('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'mcp-client', version: '1.0.0' }
    });
    proc.stdin.write(initReq);
    const initResp = await waitForResponse(10000);
    
    // 发送 initialized 通知
    proc.stdin.write(JSON.stringify({
      jsonrpc: '2.0',
      method: 'notifications/initialized',
      params: {}
    }) + '\n');
    
    initialized = true;
    console.error(`[MCP] ${serverName} 连接成功`);

    // === 步骤 2: 查看协议版本和 capabilities ===
    const serverCap = initResp.result || {};
    const protocolVer = serverCap.protocolVersion || 'unknown';

    if (action === '--list') {
      // === 列出工具 ===
      const listReq = createRequest('tools/list');
      proc.stdin.write(listReq);
      const listResp = await waitForResponse(10000);

      if (listResp.error) {
        console.error(`[错误] ${listResp.error.message}`);
        process.exit(1);
      }

      const tools = listResp.result?.tools || [];
      console.log(JSON.stringify({ server: serverName, protocolVersion: protocolVer, tools }, null, 2));
    } else {
      // === 调用工具 ===
      const toolName = action;

      // 解析参数
      const params = {};
      for (const arg of toolArgs) {
        const eqIdx = arg.indexOf('=');
        if (eqIdx > 0) {
          let key = arg.substring(0, eqIdx);
          if (key.startsWith('--')) key = key.substring(2);
          let value = arg.substring(eqIdx + 1);
          // 尝试解析 JSON 值
          try { value = JSON.parse(value); } catch (e) { /* 保留字符串 */ }
          params[key] = value;
        }
      }

      const callReq = createRequest('tools/call', {
        name: toolName,
        arguments: params
      });
      proc.stdin.write(callReq);
      const callResp = await waitForResponse(60000);

      if (callResp.error) {
        console.error(`[错误] ${callResp.error.message}`);
        process.exit(1);
      }

      console.log(JSON.stringify({
        server: serverName,
        tool: toolName,
        result: callResp.result
      }, null, 2));
    }

  } catch (err) {
    console.error(`[MCP 错误] ${err.message}`);
    process.exit(1);
  } finally {
    // 清理
    if (responseTimer) clearTimeout(responseTimer);
    proc.stdin.end();
    setTimeout(() => { proc.kill(); }, 500);
  }
}

main().catch(err => {
  console.error(err.message);
  process.exit(1);
});
