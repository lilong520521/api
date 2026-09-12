// ========================================================================
// 【改动 1】node-imap → imapflow
// ========================================================================
const { ImapFlow } = require('imapflow');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数 =====================
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
    '收件箱': 'inbox',
    'inbox': 'inbox',
    '已发送': 'sentitems',
    'sentitems': 'sentitems',
    '草稿': 'draft',
    'drafts': 'draft',
    '删除邮件': 'deleteditems',
    'deleteditems': 'deleteditems',
    '垃圾邮件': 'junkemail',
    'junk': 'junkemail'
  },
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  TARGET_FOLDERS: {
    graph: ['inbox', 'junkemail'],
    imap: ['INBOX', 'Junk'],
    chineseName: {
      'inbox': '收件箱',
      'junkemail': '垃圾箱',
      'INBOX': '收件箱',
      'Junk': '垃圾箱'
    }
  }
};

// 请求超时封装
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

// HTML特殊字符转义
function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// JSON响应特殊字符转义
function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

// 对比两封邮件，返回最新的一封
function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

// 参数校验
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
  const normalized = cleanSpecialChars.replace(/\s+/g, ' ').trim();
  return normalized.toLowerCase();
}

const VERIFY_CODE_RULES = [
  {
    regex: /(verify code|validation code|auth code|security code)[:：\s]*([0-9]{6})/i,
    desc: "英文语义6位数字验证码",
    extractGroup: 2,
    confidence: 100
  },
  {
    regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|动态口令|登录口令|验证码)[:：\s]*([0-9]{6})/i,
    desc: "中文语义6位数字验证码",
    extractGroup: 2,
    confidence: 100
  },
  {
    regex: /(验证码|校验码)[:：\s]*([0-9]{3}[-_][0-9]{3})/i,
    desc: "带分隔符的6位数字验证码",
    extractGroup: 2,
    confidence: 95
  },
  {
    regex: /\b[0-9]{6}\b/,
    desc: "纯6位数字验证码",
    extractGroup: 0,
    confidence: 90
  },
  {
    regex: /(验证码|校验码)[:：\s]*([0-9]{4})/i,
    desc: "中文语义4位数字验证码（兜底）",
    extractGroup: 2,
    confidence: 10
  },
  {
    regex: /\b[0-9]{4}\b/,
    desc: "纯4位数字验证码（最终兜底）",
    extractGroup: 0,
    confidence: 5
  }
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

  if (matchedResults.length === 0) {
    return { code: '', rule: '无匹配规则', confidence: 0 };
  }

  const uniqueResults = Array.from(new Map(matchedResults.map(item => [item.code, item])).values());
  uniqueResults.sort((a, b) => b.confidence - a.confidence);
  return uniqueResults[0];
}

function extractVerifyCodeWithLog(text, emailSubject = '未知主题') {
  const result = extractVerifyCode(text);
  console.log(`【6位验证码提取】邮件主题：${emailSubject} | 验证码：${result.code} | 匹配规则：${result.rule} | 置信度：${result.confidence}`);
  if (result.confidence < 90 && result.code) {
    console.log(`【低置信度提醒】文本片段：${preprocessText(text).substring(0, 200)}`);
  }
  return result;
}

function getVerifyCodeFromEmail(emailData, emailSubject = '未知主题') {
  const targetText = emailData.text || emailData.html || '';
  return extractVerifyCodeWithLog(targetText, emailSubject);
}

// ===================== HTML 生成 =====================
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, folderSource, verifyCode } = emailData;
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
  const folderCN = folderSource || '未知文件夹';
  const codeDisplay = verifyCode.code
    ? `<span style="color: #e53e3e; font-weight: bold; font-size: 1.2em;">${verifyCode.code}</span>（匹配规则：${verifyCode.rule}，置信度：${verifyCode.confidence}%）`
    : '未提取到6位验证码';

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${escapeHtml(subject || '无主题邮件')}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; margin: 0; padding: 20px; background: #f5f5f5; }
          .email-container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .email-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #eee; }
          .email-title { margin: 0 0 15px; color: #2d3748; }
          .email-meta { color: #4a5568; font-size: 0.9em; }
          .email-meta span { display: block; margin-bottom: 5px; }
          .email-content { color: #1a202c; }
          .email-text { white-space: pre-line; }
          .folder-source { color: #718096; font-style: italic; }
          .verify-code { margin-top: 10px; padding: 10px; background: #fef7fb; border-left: 3px solid #e53e3e; }
        </style>
      </head>
      <body>
        <div class="email-container">
          <div class="email-header">
            <h1 class="email-title">${escapeHtml(subject || '无主题')}</h1>
            <div class="email-meta">
              <span><strong>发件人：</strong>${escapeHtml(send || '未知发件人')}</span>
              <span><strong>发送日期：</strong>${new Date(date).toLocaleString() || '未知日期'}</span>
              <span class="folder-source"><strong>来源文件夹：</strong>${escapeHtml(folderCN)}</span>
              <div class="verify-code"><strong>提取的6位验证码：</strong>${codeDisplay}</div>
            </div>
          </div>
          <div class="email-content">
            ${escapedHtml}
          </div>
        </div>
      </body>
    </html>
  `;
}

// ===================== 令牌获取 =====================
// IMAP token：不传 scope（与 Go 版一致）
async function get_access_token(refresh_token, client_id) {
  try {
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

    const responseText = await response.text();
    const data = JSON.parse(responseText);
    return data.access_token;
  } catch (error) {
    throw new Error(`获取access_token失败：${error.message}`);
  }
}

// ========================================================================
// 【改动 2】graph_api 对齐 Go 版 probeGraph：
// 同时接受 Mail.Read / Mail.ReadWrite，scope 为空也放过
// 返回 { access_token, status, error }
// ========================================================================
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
      const errorText = await response.text();
      return {
        access_token: '',
        status: false,
        error: `Graph token 请求失败：${response.status} ${errorText}`
      };
    }

    const responseText = await response.text();
    const data = JSON.parse(responseText);
    const scopeStr = data.scope || '';

    // 与 Go 版一致：scope 为空 或 含 Mail.Read/Mail.ReadWrite 都视为可用
    const hasMailPermission =
      scopeStr === '' ||
      scopeStr.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1 ||
      scopeStr.indexOf('https://graph.microsoft.com/Mail.Read') !== -1;

    return {
      access_token: data.access_token || '',
      status: hasMailPermission,
      error: hasMailPermission ? '' : `Graph token 未包含 Mail 权限, scope=${scopeStr}`
    };
  } catch (error) {
    return {
      access_token: '',
      status: false,
      error: error.message || String(error)
    };
  }
}

// ===================== Graph 取件 =====================
async function get_single_folder_email(access_token, mailbox) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        "Authorization": `Bearer ${access_token}`
      },
    });

    if (!response.ok) {
      console.warn(`文件夹${mailbox}访问失败，状态码：${response.status}`);
      return null;
    }

    const responseData = await response.json();
    const email = responseData.value?.[0];
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
    console.error(`获取${mailbox}邮件失败：`, error);
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
// 【改动 3】用 imapflow 替换 node-imap
// 参数改成 { email, accessToken }，返回 { email, error, diagnostics }
// ========================================================================
async function get_dual_folder_latest_email_imap({ email, accessToken }) {
  const result = {
    email: null,
    error: '',
    diagnostics: { inboxCount: 0, junkCount: 0 }
  };

  let client = null;
  let inboxEmail = null;
  let junkEmail = null;

  try {
    client = new ImapFlow({
      host: 'outlook.office365.com',
      port: 993,
      secure: true,
      auth: {
        user: email,
        accessToken: accessToken
      },
      logger: false,                    // 👈 关键：完全禁用日志
      tls: { rejectUnauthorized: false },
      connectionTimeout: 15000,
      greetingTimeout: 10000,
      socketTimeout: 20000
    });

    // 二次保险：屏蔽残留 error 事件里的 logging 错误
    try {
      const origEmit = client.emit.bind(client);
      client.emit = function (event, ...args) {
        if (event === 'error' && args[0]) {
          const msg = args[0] && args[0].message ? args[0].message : String(args[0]);
          if (msg.includes('Logging is disabled')) {
            result.error = result.error || `日志错误（非业务）: ${msg}`;
            return false;
          }
        }
        return origEmit(event, ...args);
      };
    } catch (e) {}

    await client.connect();

    // ---- 收件箱 ----
    try {
      const lock = await client.getMailboxLock('INBOX');
      try {
        const total = client.mailbox.exists || 0;
        result.diagnostics.inboxCount = total;
        if (total > 0) {
          for await (const msg of client.fetch(
            { seq: String(total) },
            { envelope: true, source: true }
          )) {
            const mail = await simpleParser(msg.source);
            const verifyCode = getVerifyCodeFromEmail(
              { text: mail.text, html: mail.html },
              mail.subject
            );
            inboxEmail = {
              send: escapeJson(mail.from?.text || '未知发件人'),
              subject: escapeJson(mail.subject || '无主题'),
              text: escapeJson(mail.text || ''),
              html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
              date: mail.date || new Date().toISOString(),
              folderSource: '收件箱',
              verifyCode
            };
          }
        }
      } finally {
        lock.release();
      }
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      if (!msg.includes('Logging is disabled')) {
        result.error = `IMAP收件箱失败: ${msg}`;
      }
    }

    // ---- 垃圾箱 ----
    try {
      const lock = await client.getMailboxLock('Junk');
      try {
        const total = client.mailbox.exists || 0;
        result.diagnostics.junkCount = total;
        if (total > 0) {
          for await (const msg of client.fetch(
            { seq: String(total) },
            { envelope: true, source: true }
          )) {
            const mail = await simpleParser(msg.source);
            const verifyCode = getVerifyCodeFromEmail(
              { text: mail.text, html: mail.html },
              mail.subject
            );
            junkEmail = {
              send: escapeJson(mail.from?.text || '未知发件人'),
              subject: escapeJson(mail.subject || '无主题'),
              text: escapeJson(mail.text || ''),
              html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
              date: mail.date || new Date().toISOString(),
              folderSource: '垃圾箱',
              verifyCode
            };
          }
        }
      } finally {
        lock.release();
      }
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      if (!msg.includes('Logging is disabled') && !result.error) {
        result.error = `IMAP垃圾箱失败: ${msg}`;
      }
    }

    try { await client.logout(); } catch (e) {}

    result.email = getLatestEmail(inboxEmail, junkEmail);
    return result;

  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    if (!msg.includes('Logging is disabled')) {
      result.error = `IMAP连接失败: ${msg}`;
    } else {
      result.error = result.error || `IMAP日志错误: ${msg}`;
    }
    try { if (client) await client.logout(); } catch (e) {}
    return result;
  }
}

// ===================== 主入口函数 =====================
module.exports = async (req, res) => {
  let step = 'init';
  try {
    step = '1.方法校验';
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    step = '2.密码校验';
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    step = '3.参数校验';
    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);

    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // ====================================================================
    // 【改动 4】降级逻辑对齐 Go 版 GetLatestEmail
    // 1) 先试 Graph
    // 2) Graph 只要没拿到邮件 → 回退 IMAP（不管 status 是 true 还是 false）
    // 3) 两边都失败 → 返回合并错误
    // ====================================================================
    let emailInfo = null;
    let graphErr = '';
    let imapError = '';
    let imapDiagnostics = null;

    // ---- 1) 先试 Graph ----
    step = '4.Graph探活';
    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    step = '5.Graph取件';
    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过");
      emailInfo = await get_dual_folder_latest_email_graph(graph_api_result.access_token);
      if (!emailInfo) {
        graphErr = 'Graph 取件为空（收件箱/垃圾箱均无邮件）';
      }
    } else {
      graphErr = graph_api_result.error || 'Graph 权限不足';
      console.log("【降级】Graph 权限不足：", graphErr);
    }

    // ---- 2) 【关键】Graph 只要没拿到邮件就回退 IMAP ----
    step = '6.IMAP回退';
    if (!emailInfo) {
      try {
        console.log("【降级】尝试 IMAP");
        const access_token = await get_access_token(refresh_token, client_id);
        const imapResult = await get_dual_folder_latest_email_imap({
          email: email,
          accessToken: access_token
        });
        emailInfo = imapResult.email;
        imapError = imapResult.error;
        imapDiagnostics = imapResult.diagnostics;
      } catch (e) {
        imapError = `IMAP 异常: ${e.message || String(e)}`;
      }
    }

    // ---- 3) 响应 ----
    step = '7.响应生成';

    // 3.1 成功（Graph 或 IMAP 拿到邮件）
    if (emailInfo) {
      if (response_type === 'html') {
        return res.status(200).send(generateEmailHtml(emailInfo));
      }
      return res.status(200).json({
        code: 200,
        message: '6位验证码提取成功',
        data: [emailInfo]
      });
    }

    // 3.2 IMAP 明确报错 → 5000，带上两边原因
    if (imapError) {
      return res.status(500).json({
        code: 5000,
        error: `Graph失败: ${graphErr} || IMAP失败: ${imapError}`,
        diagnostics: imapDiagnostics
      });
    }

    // 3.3 IMAP 没报错但没邮件 → 真的无邮件
    return res.status(200).json({
      code: 2001,
      message: '收件箱和垃圾箱均无邮件',
      data: null,
      diagnostics: imapDiagnostics
    });

  } catch (error) {
    const msg = error && error.message ? error.message : String(error);
    let statusCode = 500;
    let errorCode = 5000;

    if (msg.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
    } else if (msg.includes('HTTP错误！状态码：403')) {
      statusCode = 403;
      errorCode = 4031;
    } else if (msg.includes('请求超时')) {
      statusCode = 504;
      errorCode = 5041;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: `步骤[${step}]失败：${msg}`
    });
  }
};
