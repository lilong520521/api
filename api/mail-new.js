// ========================================================================
// 【第一段：必须在所有 require 之前】屏蔽所有输出
// ========================================================================
(function silenceAllOutput() {
  const noop = () => {};
  const noopWrite = () => true;

  // 1. 屏蔽 console.*
  try {
    const methods = ['log','error','warn','info','debug','trace','dir','dirxml',
                     'table','time','timeEnd','group','groupEnd','assert','count',
                     'countReset','clear'];
    methods.forEach(m => {
      try { console[m] = noop; } catch (e) {
        try { Object.defineProperty(console, m, { value: noop, writable: true, configurable: true }); } catch (e2) {}
      }
    });
  } catch (e) {}

  // 2. 屏蔽 process.stdout / stderr 的 write
  try { process.stdout.write = noopWrite; } catch (e) {}
  try { process.stderr.write = noopWrite; } catch (e) {}

  // 3. 屏蔽 process.emitWarning
  try { process.emitWarning = noop; } catch (e) {}

  // 4. 屏蔽 fs.writeSync 对 fd 1/2 的写入
  try {
    const fs = require('fs');
    const origWriteSync = fs.writeSync;
    fs.writeSync = function(fd, ...args) {
      if (fd === 1 || fd === 2) return 0;
      return origWriteSync.apply(this, [fd, ...args]);
    };
  } catch (e) {}

  // 5. 屏蔽 util.debuglog
  try {
    const util = require('util');
    if (util && util.debuglog) util.debuglog = () => noop;
  } catch (e) {}
})();

// ========================================================================
// 【第二段：全局错误兜底】
// ========================================================================
(function setupGlobalHandler() {
  const isLoggingError = (e) => {
    if (!e) return false;
    const msg = (e && e.message) ? e.message : String(e);
    return msg.indexOf('Logging is disabled') !== -1 ||
           msg.indexOf('logging is disabled') !== -1 ||
           (msg.indexOf('Logging') !== -1 && msg.indexOf('disabled') !== -1);
  };

  // 供业务代码共享
  global.__isLoggingError = isLoggingError;

  try {
    process.on('uncaughtException', (err) => {
      if (isLoggingError(err)) return;
    });
  } catch (e) {}

  try {
    process.on('unhandledRejection', (reason) => {
      if (isLoggingError(reason)) return;
    });
  } catch (e) {}
})();

// ========================================================================

const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// 复用全局判断函数
const isLoggingError = global.__isLoggingError || ((e) => {
  if (!e) return false;
  const msg = (e && e.message) ? e.message : String(e);
  return msg.indexOf('Logging') !== -1 && msg.indexOf('disabled') !== -1;
});

// ===================== 安全日志 =====================
function safeLog(...args) { try { console.log(...args); } catch (e) {} }
function safeError(...args) { try { console.error(...args); } catch (e) {} }

// ===================== 全局配置 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    tlsOptions: { rejectUnauthorized: false },
    connTimeout: 10000,
    authTimeout: 10000
  },
  MAILBOX_MAP: {
    '收件箱': 'inbox', 'inbox': 'inbox',
    '已发送': 'sentitems', 'sentitems': 'sentitems',
    '草稿': 'draft', 'drafts': 'draft',
    '删除邮件': 'deleteditems', 'deleteditems': 'deleteditems',
    '垃圾邮件': 'junkemail', 'junk': 'junkemail'
  },
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  TARGET_FOLDERS: {
    graph: ['inbox', 'junkemail'],
    imap: ['INBOX', 'Junk'],
    chineseName: {
      'inbox': '收件箱', 'junkemail': '垃圾箱',
      'INBOX': '收件箱', 'Junk': '垃圾箱'
    }
  }
};

// ===================== 工具函数 =====================
async function fetchWithTimeout(url, options = {}, timeout = CONFIG.REQUEST_TIMEOUT) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    clearTimeout(timeoutId);
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    throw new Error(error.name === "AbortError" ? "请求超时（超过10秒）" : error.message);
  }
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 验证码提取 =====================
function preprocessText(rawText) {
  if (!rawText) return '';
  const withoutHtml = rawText.replace(/<[^>]+>/g, '');
  const mergeDigitSeparators = withoutHtml.replace(/(\d)[\s-_]+(\d)/g, '$1$2');
  const cleanSpecialChars = mergeDigitSeparators.replace(/[^\u4e00-\u9fa5a-zA-Z0-9，。：！？]/g, '');
  return cleanSpecialChars.replace(/\s+/g, ' ').trim().toLowerCase();
}

const VERIFY_CODE_RULES = [
  { regex: /(verify code|validation code|auth code|security code)[:：\s]*([0-9]{6})/i, desc: "英文语义6位", extractGroup: 2, confidence: 100 },
  { regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|动态口令|登录口令)[:：\s]*([0-9]{6})/i, desc: "中文语义6位", extractGroup: 2, confidence: 100 },
  { regex: /(验证码|校验码)[:：\s]*([0-9]{3}[-_][0-9]{3})/i, desc: "带分隔符6位", extractGroup: 2, confidence: 95 },
  { regex: /\b[0-9]{6}\b/, desc: "纯6位", extractGroup: 0, confidence: 90 },
  { regex: /(验证码|校验码)[:：\s]*([0-9]{4})/i, desc: "中文语义4位兜底", extractGroup: 2, confidence: 10 },
  { regex: /\b[0-9]{4}\b/, desc: "纯4位兜底", extractGroup: 0, confidence: 5 }
];

function extractVerifyCode(text) {
  const cleanText = preprocessText(text);
  if (!cleanText) return { code: '', rule: '无有效文本', confidence: 0 };
  const matchedResults = [];
  for (const rule of VERIFY_CODE_RULES) {
    const matches = cleanText.match(rule.regex);
    if (matches) {
      const code = matches[rule.extractGroup].trim();
      const isSixDigit = code.length === 6 && /^\d{6}$/.test(code);
      if (isSixDigit) {
        matchedResults.push({ code, rule: rule.desc, confidence: rule.confidence });
      } else if (!isSixDigit && rule.confidence < 90) {
        matchedResults.push({ code, rule: rule.desc, confidence: rule.confidence });
      }
    }
  }
  if (matchedResults.length === 0) return { code: '', rule: '无匹配规则', confidence: 0 };
  const uniqueResults = Array.from(new Map(matchedResults.map(item => [item.code, item])).values());
  uniqueResults.sort((a, b) => b.confidence - a.confidence);
  return uniqueResults[0];
}

function getVerifyCodeFromEmail(emailData, emailSubject = '未知主题') {
  const targetText = emailData.text || emailData.html || '';
  const result = extractVerifyCode(targetText);
  safeLog(`【验证码】主题：${emailSubject} | 验证码：${result.code}`);
  return result;
}

// ===================== HTML 生成 =====================
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, folderSource, verifyCode } = emailData || {};
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
  const folderCN = folderSource || '未知文件夹';
  const codeDisplay = verifyCode && verifyCode.code
    ? `<span style="color: #e53e3e; font-weight: bold; font-size: 1.2em;">${verifyCode.code}</span>`
    : '未提取到6位验证码';

  return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>${escapeHtml(subject || '无主题')}</title>
    <style>body{font-family:-apple-system,sans-serif;line-height:1.6;padding:20px;background:#f5f5f5}
    .c{max-width:800px;margin:0 auto;background:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.1)}
    .h{margin-bottom:20px;padding-bottom:15px;border-bottom:1px solid #eee}
    .vc{margin-top:10px;padding:10px;background:#fef7fb;border-left:3px solid #e53e3e}</style></head>
    <body><div class="c"><div class="h"><h1>${escapeHtml(subject || '无主题')}</h1>
    <p>发件人：${escapeHtml(send || '未知')}</p>
    <p>日期：${date ? new Date(date).toLocaleString() : '未知'}</p>
    <p>来源：${escapeHtml(folderCN)}</p>
    <div class="vc"><strong>验证码：</strong>${codeDisplay}</div></div>
    <div>${escapedHtml}</div></div></body></html>`;
}

// ===================== 核心业务 =====================
async function get_access_token(refresh_token, client_id) {
  const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      'client_id': client_id,
      'grant_type': 'refresh_token',
      'refresh_token': refresh_token
    }).toString()
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
  }
  const data = await response.json();
  return data.access_token;
}

const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

async function graph_api(refresh_token, client_id) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/.default'
      }).toString()
    });
    if (!response.ok) {
      return { access_token: '', status: false, error: `Graph token 请求失败：${response.status}` };
    }
    const data = await response.json();
    const scopeStr = data.scope || '';
    const hasMailPermission =
      scopeStr === '' ||
      scopeStr.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1 ||
      scopeStr.indexOf('https://graph.microsoft.com/Mail.Read') !== -1;
    return {
      access_token: data.access_token || '',
      status: hasMailPermission,
      error: hasMailPermission ? '' : `scope 未含 Graph Mail 权限`
    };
  } catch (error) {
    return { access_token: '', status: false, error: error.message || String(error) };
  }
}

async function get_single_folder_email(access_token, mailbox) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json', "Authorization": `Bearer ${access_token}` }
    });
    if (!response.ok) return null;
    const data = await response.json();
    const email = data.value?.[0];
    if (!email) return null;
    const verifyCode = getVerifyCodeFromEmail(
      { text: email['bodyPreview'], html: email['body']?.['content'] },
      email['subject']
    );
    return {
      send: email['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: email['subject'] || '无主题',
      text: email['bodyPreview'] || '',
      html: email['body']?.['content'] || '',
      date: email['createdDateTime'] || new Date().toISOString(),
      folderSource: CONFIG.TARGET_FOLDERS.chineseName[mailbox] || '未知文件夹',
      verifyCode
    };
  } catch (error) {
    return null;
  }
}

async function get_dual_folder_latest_email_graph(access_token) {
  const [inboxEmail, junkEmail] = await Promise.all([
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[0]),
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[1])
  ]);
  return getLatestEmail(inboxEmail, junkEmail);
}

// ========================================================================
// 【核心修复】IMAP 流程：返回 { email, error, diagnostics }
// ========================================================================
async function get_dual_folder_latest_email_imap(imapConfig) {
  return new Promise((resolve) => {
    let imap = null;
    let resolved = false;
    let lastError = '';
    let inboxCount = 0;
    let junkCount = 0;

    const safeResolve = (value) => {
      if (resolved) return;
      resolved = true;
      try { if (imap) imap.end(); } catch (e) {}
      resolve({
        email: value,
        error: lastError,
        diagnostics: { inboxCount, junkCount }
      });
    };

    try {
      imap = new Imap(imapConfig);

      const origEmit = imap.emit.bind(imap);
      imap.emit = function(event, ...args) {
        if (event === 'error' && args[0]) {
          const msg = args[0] && args[0].message ? args[0].message : String(args[0]);
          if (isLoggingError(args[0])) {
            lastError = lastError || `imap日志错误（非业务）: ${msg}`;
            return false;
          }
        }
        return origEmit(event, ...args);
      };

      imap.once('ready', async () => {
        let inboxEmail = null;
        let junkEmail = null;

        // ---- 收件箱 ----
        try {
          const inboxFolder = CONFIG.TARGET_FOLDERS.imap[0];
          await new Promise((res, rej) => {
            imap.openBox(inboxFolder, true, (err) => {
              if (err) {
                if (isLoggingError(err)) { res(); return; }
                rej(err); return;
              }
              res();
            });
          });

          const inboxResults = await new Promise((res, rej) => {
            imap.search(["ALL"], (err, resArr) => {
              if (err) {
                if (isLoggingError(err)) { res([]); return; }
                rej(err); return;
              }
              res(resArr);
            });
          });

          inboxCount = inboxResults ? inboxResults.length : 0;

          if (inboxResults && inboxResults.length > 0) {
            const latestInbox = inboxResults.slice(-1);
            const f1 = imap.fetch(latestInbox, { bodies: "" });
            await new Promise((res) => {
              f1.on('message', async (msg) => {
                try {
                  const stream = await new Promise((r) => msg.on("body", r));
                  const mail = await simpleParser(stream);
                  const verifyCode = getVerifyCodeFromEmail(
                    { text: mail.text, html: mail.html }, mail.subject
                  );
                  inboxEmail = {
                    send: escapeJson(mail.from?.text || '未知发件人'),
                    subject: escapeJson(mail.subject || '无主题'),
                    text: escapeJson(mail.text || ''),
                    html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                    date: mail.date || new Date().toISOString(),
                    folderSource: CONFIG.TARGET_FOLDERS.chineseName[inboxFolder] || '未知文件夹',
                    verifyCode
                  };
                } catch (e) {
                  lastError = `解析收件箱邮件失败: ${e.message}`;
                }
                res();
              });
              f1.once('error', (e) => {
                if (!isLoggingError(e)) lastError = `fetch收件箱失败: ${e.message}`;
                res();
              });
            });
          }
        } catch (err) {
          if (!isLoggingError(err)) {
            lastError = `IMAP收件箱失败: ${err.message}`;
          }
        }

        // ---- 垃圾箱 ----
        try {
          const junkFolder = CONFIG.TARGET_FOLDERS.imap[1];
          await new Promise((res, rej) => {
            imap.openBox(junkFolder, true, (err) => {
              if (err) {
                if (isLoggingError(err)) { res(); return; }
                rej(err); return;
              }
              res();
            });
          });

          const junkResults = await new Promise((res, rej) => {
            imap.search(["ALL"], (err, resArr) => {
              if (err) {
                if (isLoggingError(err)) { res([]); return; }
                rej(err); return;
              }
              res(resArr);
            });
          });

          junkCount = junkResults ? junkResults.length : 0;

          if (junkResults && junkResults.length > 0) {
            const latestJunk = junkResults.slice(-1);
            const f2 = imap.fetch(latestJunk, { bodies: "" });
            await new Promise((res) => {
              f2.on('message', async (msg) => {
                try {
                  const stream = await new Promise((r) => msg.on("body", r));
                  const mail = await simpleParser(stream);
                  const verifyCode = getVerifyCodeFromEmail(
                    { text: mail.text, html: mail.html }, mail.subject
                  );
                  junkEmail = {
                    send: escapeJson(mail.from?.text || '未知发件人'),
                    subject: escapeJson(mail.subject || '无主题'),
                    text: escapeJson(mail.text || ''),
                    html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                    date: mail.date || new Date().toISOString(),
                    folderSource: CONFIG.TARGET_FOLDERS.chineseName[junkFolder] || '未知文件夹',
                    verifyCode
                  };
                } catch (e) {
                  lastError = `解析垃圾箱邮件失败: ${e.message}`;
                }
                res();
              });
              f2.once('error', (e) => {
                if (!isLoggingError(e)) lastError = `fetch垃圾箱失败: ${e.message}`;
                res();
              });
            });
          }
        } catch (err) {
          if (!isLoggingError(err)) {
            lastError = lastError || `IMAP垃圾箱失败: ${err.message}`;
          }
        }

        safeResolve(getLatestEmail(inboxEmail, junkEmail));
      });

      imap.once('error', (err) => {
        if (isLoggingError(err)) {
          lastError = lastError || `IMAP日志错误: ${err.message}`;
        } else {
          lastError = `IMAP连接错误: ${err.message}`;
        }
        safeResolve(null);
      });

      try {
        imap.connect();
      } catch (e) {
        if (isLoggingError(e)) {
          lastError = lastError || `imap.connect日志错误: ${e.message}`;
        } else {
          lastError = `imap.connect异常: ${e.message}`;
        }
        safeResolve(null);
      }

      setTimeout(() => {
        if (!resolved) {
          lastError = lastError || 'IMAP 20秒超时';
          safeResolve(null);
        }
      }, 20000);

    } catch (e) {
      lastError = `IMAP外层异常: ${e.message}`;
      safeResolve(null);
    }
  });
}

// ===================== 主入口 =====================
module.exports = async (req, res) => {
  try {
    const noop = () => {};
    try { process.stdout.write = function() { return true; }; } catch (e) {}
    try { process.stderr.write = function() { return true; }; } catch (e) {}
    try { ['log','error','warn','info','debug','trace','dir','table','assert'].forEach(m => { console[m] = noop; }); } catch (e) {}
  } catch (e) {}

  let step = 'init';
  try {
    step = '1.方法校验';
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({ code: 405, error: `不支持的请求方法` });
    }

    step = '2.密码校验';
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({ code: 4010, error: '认证失败' });
    }

    step = '3.参数校验';
    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      return res.status(400).json({ code: 4001, error: `缺少必要参数：${missingParams.join('、')}` });
    }
    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({ code: 4002, error: paramError.message });
    }

    step = '4.Graph探活';
    const graph_api_result = await graph_api(refresh_token, client_id);
    let emailInfo = null;
    let graphErr = '';

    step = '5.Graph取件';
    if (graph_api_result.status) {
      emailInfo = await get_dual_folder_latest_email_graph(graph_api_result.access_token);
      if (!emailInfo) graphErr = 'Graph 取件为空';
    } else {
      graphErr = graph_api_result.error || 'Graph 权限不足';
    }

    step = '6.IMAP回退';
    let imapError = '';
    let imapDiagnostics = null;
    if (!emailInfo) {
      const access_token = await get_access_token(refresh_token, client_id);
      const authString = generateAuthString(email, access_token);
      const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };
      const imapResult = await get_dual_folder_latest_email_imap(imapConfig);
      emailInfo = imapResult.email;
      imapError = imapResult.error;
      imapDiagnostics = imapResult.diagnostics;
    }

    step = '7.响应生成';

    // ====================================================================
    // 【关键】区分三种情况：真无邮件 / IMAP出错 / 成功
    // ====================================================================
    if (!emailInfo) {
      // 情况1：IMAP 出错
      if (imapError) {
        if (response_type === 'html') {
          return res.status(500).send(
            `<html><body><h1>IMAP 出错</h1><p>${escapeHtml(imapError)}</p>` +
            `<p>诊断：收件箱 ${imapDiagnostics?.inboxCount || 0} 封，垃圾箱 ${imapDiagnostics?.junkCount || 0} 封</p></body></html>`
          );
        }
        return res.status(500).json({
          code: 5000,
          error: `Graph失败: ${graphErr} || IMAP失败: ${imapError}`,
          diagnostics: imapDiagnostics
        });
      }

      // 情况2：有邮件但 fetch 失败（诊断有邮件数但没拿到内容）
      if (imapDiagnostics && (imapDiagnostics.inboxCount > 0 || imapDiagnostics.junkCount > 0)) {
        if (response_type === 'html') {
          return res.status(500).send(`<html><body><h1>有邮件但解析失败</h1></body></html>`);
        }
        return res.status(500).json({
          code: 5000,
          error: `IMAP有邮件但解析失败`,
          diagnostics: imapDiagnostics
        });
      }

      // 情况3：真的无邮件
      if (response_type === 'html') {
        return res.status(200).send(generateEmailHtml({}));
      }
      return res.status(200).json({
        code: 2001,
        message: "收件箱和垃圾箱确实均无邮件",
        data: null,
        diagnostics: imapDiagnostics
      });
    }

    // 成功
    if (response_type === 'html') {
      return res.status(200).send(generateEmailHtml(emailInfo));
    }
    return res.status(200).json({
      code: 200,
      message: '6位验证码提取成功',
      data: [emailInfo]
    });

  } catch (error) {
    const msg = error && error.message ? error.message : String(error);

    // 如果是 logging 类错误，返回"无邮件"而不是 500
    if (isLoggingError(error)) {
      return res.status(200).json({
        code: 2001,
        message: "IMAP 流程被 Vercel 日志拦截，但业务已完成（返回空）",
        data: null
      });
    }

    let statusCode = 500, errorCode = 5000;
    if (msg.includes('HTTP错误！状态码：401')) { statusCode = 401; errorCode = 4011; }
    else if (msg.includes('HTTP错误！状态码：403')) { statusCode = 403; errorCode = 4031; }
    else if (msg.includes('请求超时')) { statusCode = 504; errorCode = 5041; }
    return res.status(statusCode).json({ code: errorCode, error: `步骤[${step}]失败：${msg}` });
  }
};
