const Imap = require('node-imap');
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

// ===================== 验证码提取（强制优先6位） =====================
// 1. 文本预处理：清洗+合并数字分隔符
function preprocessText(rawText) {
  if (!rawText) return '';

  // 移除HTML标签
  const withoutHtml = rawText.replace(/<[^>]+>/g, '');
  // 合并数字中的空格/分隔符（-/_/空格）→ 确保6位数字连续
  const mergeDigitSeparators = withoutHtml.replace(/(\d)[\s-_]+(\d)/g, '$1$2');
  // 保留数字、字母、中文、常用标点，移除其他特殊字符
  const cleanSpecialChars = mergeDigitSeparators.replace(/[^\u4e00-\u9fa5a-zA-Z0-9，。：！？]/g, '');
  // 去除多余空格/换行，统一为单个空格
  const normalized = cleanSpecialChars.replace(/\s+/g, ' ').trim();
  // 统一大小写
  return normalized.toLowerCase();
}

// 2. 验证码规则库（仅保留6位相关，4位仅作为最终兜底）
const VERIFY_CODE_RULES = [
  // 最高优先级：英文语义6位数字
  {
    regex: /(verify code|validation code|auth code|security code)[:：\s]*([0-9]{6})/i,
    desc: "英文语义6位数字验证码",
    extractGroup: 2,
    confidence: 100
  },
  // 次高优先级：中文语义6位数字（覆盖所有常见关键词）
  {
    regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|动态口令|登录口令|验证码)[:：\s]*([0-9]{6})/i,
    desc: "中文语义6位数字验证码",
    extractGroup: 2,
    confidence: 100
  },
  // 第三优先级：带分隔符的6位数字（如123-456 → 预处理后已合并为123456）
  {
    regex: /(验证码|校验码)[:：\s]*([0-9]{3}[-_][0-9]{3})/i,
    desc: "带分隔符的6位数字验证码",
    extractGroup: 2,
    confidence: 95
  },
  // 第四优先级：纯6位数字（兜底，无语义也优先6位）
  {
    regex: /\b[0-9]{6}\b/,
    desc: "纯6位数字验证码",
    extractGroup: 0,
    confidence: 90
  },
  // 最低优先级：4位（仅无6位时返回）
  {
    regex: /(验证码|校验码)[:：\s]*([0-9]{4})/i,
    desc: "中文语义4位数字验证码（兜底）",
    extractGroup: 2,
    confidence: 10
  },
  // 最终兜底：纯4位数字（仅无任何6位时返回）
  {
    regex: /\b[0-9]{4}\b/,
    desc: "纯4位数字验证码（最终兜底）",
    extractGroup: 0,
    confidence: 5
  }
];

// 3. 核心提取函数（强制优先6位）
function extractVerifyCode(text) {
  const cleanText = preprocessText(text);
  if (!cleanText) return { code: '', rule: '无有效文本', confidence: 0 };

  // 遍历规则库
  const matchedResults = [];
  for (const rule of VERIFY_CODE_RULES) {
    const matches = cleanText.match(rule.regex);
    if (matches) {
      const code = matches[rule.extractGroup].trim();
      // 确保6位数字完整性（过滤误匹配的非6位）
      const isSixDigit = code.length === 6 && /^\d{6}$/.test(code);
      if (isSixDigit) {
        matchedResults.push({
          code,
          rule: rule.desc,
          confidence: rule.confidence
        });
      } else if (!isSixDigit && rule.confidence < 90) { // 仅4位规则允许非6位
        matchedResults.push({
          code,
          rule: rule.desc,
          confidence: rule.confidence
        });
      }
    }
  }

  // 无匹配结果
  if (matchedResults.length === 0) {
    return { code: '', rule: '无匹配规则', confidence: 0 };
  }

  // 去重 + 按置信度排序（6位始终优先）
  const uniqueResults = Array.from(new Map(matchedResults.map(item => [item.code, item])).values());
  uniqueResults.sort((a, b) => b.confidence - a.confidence);
  
  return uniqueResults[0];
}

// 4. 带日志的提取函数
function extractVerifyCodeWithLog(text, emailSubject = '未知主题') {
  const result = extractVerifyCode(text);
  console.log(`【6位验证码提取】邮件主题：${emailSubject} | 验证码：${result.code} | 匹配规则：${result.rule} | 置信度：${result.confidence}`);
  // 低置信度（<90）时打印文本片段，便于调试
  if (result.confidence < 90 && result.code) {
    console.log(`【低置信度提醒】文本片段：${preprocessText(text).substring(0, 200)}`);
  }
  return result;
}

// 5. 从邮件数据提取验证码
function getVerifyCodeFromEmail(emailData, emailSubject = '未知主题') {
  const targetText = emailData.text || emailData.html || '';
  return extractVerifyCodeWithLog(targetText, emailSubject);
}

// ===================== 核心业务函数 =====================
// 生成邮件HTML（含6位验证码高亮）
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, folderSource, verifyCode } = emailData;
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
  const folderCN = folderSource || '未知文件夹';
  // 6位验证码高亮显示
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

// 获取access_token
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

// 生成IMAP认证字符串
const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

// 检查Graph API权限
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
      throw new Error(`Graph API请求失败：状态码${response.status}，响应：${errorText}`);
    }

    const responseText = await response.text();
    const data = JSON.parse(responseText);
    const hasMailPermission = data.scope?.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;

    return {
      access_token: data.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

// 单个文件夹取件（Graph API）
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

    // 提取6位验证码
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

// Graph API双文件夹取最新邮件
async function get_dual_folder_latest_email_graph(access_token) {
  const [inboxEmail, junkEmail] = await Promise.all([
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[0]),
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[1])
  ]);
  return getLatestEmail(inboxEmail, junkEmail);
}

// IMAP双文件夹取最新邮件
async function get_dual_folder_latest_email_imap(imapConfig) {
  const imap = new Imap(imapConfig);
  let inboxEmail = null;
  let junkEmail = null;

  const fetchEmails = new Promise((resolve, reject) => {
    imap.once('ready', async () => {
      try {
        // 1. 获取收件箱邮件
        try {
          const inboxFolder = CONFIG.TARGET_FOLDERS.imap[0];
          await new Promise((res, rej) => {
            imap.openBox(inboxFolder, true, (err) => err ? rej(err) : res());
          });
          const inboxResults = await new Promise((res, rej) => {
            imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr));
          });
          if (inboxResults.length > 0) {
            const latestInbox = inboxResults.slice(-1);
            const f1 = imap.fetch(latestInbox, { bodies: "" });
            await new Promise((res) => {
              f1.on('message', async (msg) => {
                const stream = await new Promise((r) => msg.on("body", r));
                const mail = await simpleParser(stream);
                // 提取6位验证码
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
                  folderSource: CONFIG.TARGET_FOLDERS.chineseName[inboxFolder] || '未知文件夹',
                  verifyCode
                };
                res();
              });
            });
          }
        } catch (err) {
          console.error('IMAP获取收件箱邮件失败：', err);
        }

        // 2. 获取垃圾箱邮件
        try {
          const junkFolder = CONFIG.TARGET_FOLDERS.imap[1];
          await new Promise((res, rej) => {
            imap.openBox(junkFolder, true, (err) => err ? rej(err) : res());
          });
          const junkResults = await new Promise((res, rej) => {
            imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr));
          });
          if (junkResults.length > 0) {
            const latestJunk = junkResults.slice(-1);
            const f2 = imap.fetch(latestJunk, { bodies: "" });
            await new Promise((res) => {
              f2.on('message', async (msg) => {
                const stream = await new Promise((r) => msg.on("body", r));
                const mail = await simpleParser(stream);
                // 提取6位验证码
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
                  folderSource: CONFIG.TARGET_FOLDERS.chineseName[junkFolder] || '未知文件夹',
                  verifyCode
                };
                res();
              });
            });
          }
        } catch (err) {
          console.error('IMAP获取垃圾箱邮件失败：', err);
        }

        imap.end();
        resolve(getLatestEmail(inboxEmail, junkEmail));
      } catch (err) {
        imap.end();
        reject(err);
      }
    });

    imap.once('error', (err) => reject(err));
    imap.connect();
  });

  return fetchEmails;
}

// ===================== 主入口函数 =====================
module.exports = async (req, res) => {
  try {
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

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

    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，获取收件箱+垃圾箱最新邮件");
      const latestEmail = await get_dual_folder_latest_email_graph(graph_api_result.access_token);

      if (!latestEmail) {
        return res.status(200).json({
          code: 2001,
          message: "收件箱和垃圾箱均无邮件",
          data: null
        });
      }

      if (response_type === 'html') {
        const htmlResponse = generateEmailHtml(latestEmail);
        return res.status(200).send(htmlResponse);
      } else {
        return res.status(200).json({
          code: 200,
          message: '6位验证码提取成功',
          data: [latestEmail]
        });
      }
    }

    console.log("【降级】Graph API权限不足，使用IMAP获取收件箱+垃圾箱最新邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };

    const latestEmailImap = await get_dual_folder_latest_email_imap(imapConfig);

    if (!latestEmailImap) {
      return res.status(200).json({
        code: 2001,
        message: "收件箱和垃圾箱均无邮件",
        data: null
      });
    }

    if (response_type === 'html') {
      const htmlResponse = generateEmailHtml(latestEmailImap);
      return res.status(200).send(htmlResponse);
    } else {
      return res.status(200).json({
        code: 200,
        message: '6位验证码提取成功',
        data: [latestEmailImap]
      });
    }

  } catch (error) {
    let statusCode = 500;
    let errorCode = 5000;

    if (error.message.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      error.message = '认证失效，请刷新refresh_token';
    } else if (error.message.includes('HTTP错误！状态码：403')) {
      statusCode = 403;
      errorCode = 4031;
      error.message = '权限不足，需开启Mail.ReadWrite权限';
    } else if (error.message.includes('请求超时')) {
      statusCode = 504;
      errorCode = 5041;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: `服务器错误：${error.message}`
    });
  }
};
