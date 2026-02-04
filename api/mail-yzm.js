const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;
const atob = require('atob');

// ===================== 全局配置（无变更） =====================
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
  },
  FILTERED_NUMBERS: {
    timeRegex: /^(0\d|1\d|2[0-3])([0-5]\d)([0-5]\d)$/,
    dateRegex: /^(20[2-9]\d)(0[1-9]|1[0-2])$/,
    repeatRegex: /^(\d)\1{5}$/,
    sequenceRegex: /^(012345|123456|234567|345678|456789|567890|987654|876543|765432|654321|543210)$/
  }
};

// ===================== 工具函数（无变更） =====================
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
  if (!emailReg.test(email)) return new Error("邮箱格式无效");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 验证码提取（已修复matchAll错误） =====================
function preprocessText(rawText) {
  if (!rawText) return '';

  let textWithAlt = rawText.replace(/<img[^>]+alt="([^"]+)"/gi, (match, alt) => ` ${alt} `);
  const withoutHtml = textWithAlt.replace(/<[^>]+>/g, ' ');
  
  const base64Regex = /(?:data:image\/\w+;base64,)?([A-Za-z0-9+/=]{10,})/g;
  textWithAlt = withoutHtml.replace(base64Regex, (match, base64Str) => {
    try {
      const decoded = atob(base64Str);
      return /^\d{6}$/.test(decoded) ? decoded : match;
    } catch (e) {
      return match;
    }
  });
  
  const mergeSeparators = textWithAlt
    .replace(/[\s\.\,\|\-\_\[\]\(\)\{\}\:：；]/g, '')
    .replace(/(\d)([a-zA-Z\u4e00-\u9fa5])(\d)/g, '$1$3');

  const contextRegex = /(.{0,20})(验证码|校验码|动态码|登录码|安全码|短信码|授权码|临时码|激活码|verify code|validation code|auth code|security code)(.{0,20})/gi;
  let contextText = '';
  let match;
  while ((match = contextRegex.exec(mergeSeparators)) !== null) {
    contextText += match[1] + match[2] + match[3] + ' ';
  }

  const targetText = contextText.trim() || mergeSeparators;
  return targetText.toLowerCase().trim();
}

const VERIFY_CODE_RULES = [
  {
    regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|临时码|激活码|verify code|validation code|auth code|security code)[:：\s]*[【\(\{]?[0-9]{6}[】\)\}]?/ig,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 100
  },
  {
    regex: /(v|code|verify|auth|激活码|验证码)[0-9]{6}/ig,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 98
  },
  {
    regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|临时码|激活码|verify code|validation code|auth code|security code).{0,10}[0-9]{6}/ig,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 95
  },
  {
    regex: /[【\(\{][0-9]{6}[】\)\}]/g,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 90
  },
  {
    regex: /[0-9]{3}[-\.\_]{1}[0-9]{3}/g,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 85
  },
  {
    regex: /\b[0-9]{6}\b/g,
    extractFn: (match) => match[0],
    confidence: 80
  },
  {
    regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|verify code|validation code|auth code|security code)[:：\s]*[【\(\{]?[0-9]{4}[】\)\}]?/ig,
    extractFn: (match) => match[0].replace(/[^0-9]/g, ''),
    confidence: 10
  },
  {
    regex: /\b[0-9]{4}\b/g,
    extractFn: (match) => match[0],
    confidence: 5
  }
];

function filterInvalidCode(code) {
  if (!code || code.length !== 6) return false;
  const { timeRegex, dateRegex, repeatRegex, sequenceRegex } = CONFIG.FILTERED_NUMBERS;
  if (timeRegex.test(code)) return true;
  if (dateRegex.test(code)) return true;
  if (repeatRegex.test(code)) return true;
  if (sequenceRegex.test(code)) return true;
  return false;
}

function extractVerifyCode(text) {
  const cleanText = preprocessText(text);
  if (!cleanText) return '';

  const matchedResults = [];
  for (const rule of VERIFY_CODE_RULES) {
    const matches = cleanText.matchAll(rule.regex);
    for (const match of matches) {
      const code = rule.extractFn(match);
      if (code.length === 6 && filterInvalidCode(code)) continue;
      if (!matchedResults.some(item => item === code)) {
        matchedResults.push(code);
      }
    }
  }

  if (matchedResults.length === 0) return '';
  const sixDigitCode = matchedResults.find(code => code.length === 6);
  return sixDigitCode || matchedResults[0];
}

function getVerifyCodeFromEmail(emailData) {
  const textSources = [
    emailData.text || '',
    emailData.html || '',
    emailData.subject || '',
    emailData.from?.text || ''
  ].join(' ');
  return extractVerifyCode(textSources);
}

// ===================== 响应生成（关键修改：黑色背景+欢迎语） =====================
// HTML响应：黑色背景窗口 + 顶部欢迎语 + 核心信息 + 验证码高亮
function generateCodeHtmlWithMeta(verifyCode, sender, sendDate, folder) {
  const codeDisplay = verifyCode || '未提取到验证码';
  const codeStyle = verifyCode && verifyCode.length === 6
    ? 'color: #e53e3e; font-weight: bold; font-size: 3em;'
    : 'color: #ccc; font-size: 2em;';
  // 格式化日期
  const formatDate = new Date(sendDate).toLocaleString() || '未知日期';
  // 格式化核心信息
  const showSender = sender || '未知发件人';
  const showFolder = folder || '未知文件夹';

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>小黑API取件系统</title>
        <style>
          /* 整体黑色背景 */
          body { 
            margin: 0; 
            padding: 0; 
            min-height: 100vh; 
            display: flex; 
            justify-content: center; 
            align-items: center; 
            background-color: #000; /* 窗口背景设为纯黑 */
          }
          /* 内容容器：居中显示，轻微透明黑底增强层次感 */
          .container { 
            text-align: center; 
            padding: 40px 30px; 
            background-color: rgba(0, 0, 0, 0.8); /* 半透黑底（可选，也可直接#000） */
            border-radius: 12px; 
            box-shadow: 0 0 20px rgba(255, 255, 255, 0.1); /* 白色微光阴影，增强黑色背景下的立体感 */
            width: 90%;
            max-width: 500px;
          }
          /* 欢迎语样式：醒目白色，大号字体，加粗，底部间距 */
          .welcome-title { 
            color: #fff; /* 白色字体 */
            font-size: 1.8em; 
            font-weight: bold; 
            margin-bottom: 30px; 
            padding-bottom: 15px;
            border-bottom: 1px solid #333; /* 灰色分隔线，区分欢迎语和内容 */
          }
          /* 元信息样式：浅灰色，清晰不喧宾夺主 */
          .meta-info { 
            margin-bottom: 25px; 
            font-size: 1em; 
            color: #ccc; /* 浅灰色字体，黑色背景下更易读 */
            line-height: 1.8; 
            text-align: left;
            padding: 0 20px;
          }
          /* 验证码样式：高亮显示，居中 */
          .code-text { 
            ${codeStyle} 
            letter-spacing: 4px; /* 增加字符间距，更易识别 */
            margin: 10px 0; 
            text-shadow: 0 0 10px rgba(229, 62, 62, 0.5); /* 6位验证码添加红色微光阴影 */
          }
        </style>
      </head>
      <body>
        <div class="container">
          <!-- 顶部欢迎语：固定内容 -->
          <div class="welcome-title">欢迎使用小黑api取件 系统</div>
          <!-- 核心元信息：发件人、日期、文件夹 -->
          <div class="meta-info">
            <p>📧 发件人：${showSender}</p>
            <p>📅 发送日期：${formatDate}</p>
            <p>📁 来源文件夹：${showFolder}</p>
          </div>
          <!-- 验证码高亮显示 -->
          <div class="code-text">${codeDisplay}</div>
        </div>
      </body>
    </html>
  `;
}

// ===================== 核心业务函数（无变更） =====================
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

    const data = await response.json();
    return data.access_token;
  } catch (error) {
    throw new Error(`获取access_token失败：${error.message}`);
  }
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
      const errorText = await response.text();
      throw new Error(`Graph API请求失败：状态码${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    const hasMailPermission = data.scope?.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;
    return { access_token: data.access_token, status: hasMailPermission };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

async function get_single_folder_email_graph(access_token, mailbox) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc&$select=from,subject,bodyPreview,body,createdDateTime`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        "Authorization": `Bearer ${access_token}`
      },
    });

    if (!response.ok) {
      console.warn(`文件夹${mailbox}访问失败`);
      return null;
    }

    const responseData = await response.json();
    const email = responseData.value?.[0];
    if (!email) return null;

    const sender = email['from']?.['emailAddress']?.['address'] || email['from']?.['emailAddress']?.['name'] || '未知发件人';
    const sendDate = email['createdDateTime'] || new Date().toISOString();
    const folder = CONFIG.TARGET_FOLDERS.chineseName[mailbox] || '未知文件夹';
    const verifyCode = getVerifyCodeFromEmail({
      text: email['bodyPreview'] || '',
      html: email['body']?.['content'] || '',
      subject: email['subject'] || '',
      from: { text: sender }
    });

    return {
      sender,
      sendDate,
      folder,
      verifyCode
    };
  } catch (error) {
    console.error(`获取${mailbox}邮件失败：`, error);
    return null;
  }
}

async function get_dual_folder_latest_email_graph(access_token) {
  const [inboxEmail, junkEmail] = await Promise.all([
    get_single_folder_email_graph(access_token, CONFIG.TARGET_FOLDERS.graph[0]),
    get_single_folder_email_graph(access_token, CONFIG.TARGET_FOLDERS.graph[1])
  ]);
  return getLatestEmail(inboxEmail, junkEmail);
}

async function get_dual_folder_latest_email_imap(imapConfig) {
  const imap = new Imap(imapConfig);
  let inboxEmail = null;
  let junkEmail = null;

  const fetchEmails = new Promise((resolve, reject) => {
    imap.once('ready', async () => {
      try {
        // 获取收件箱邮件及信息
        try {
          const inboxFolder = CONFIG.TARGET_FOLDERS.imap[0];
          await new Promise((res, rej) => imap.openBox(inboxFolder, true, (err) => err ? rej(err) : res()));
          const inboxResults = await new Promise((res, rej) => imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr)));
          if (inboxResults.length > 0) {
            const latestInbox = inboxResults.slice(-1);
            const f1 = imap.fetch(latestInbox, { bodies: "" });
            await new Promise((res) => {
              f1.on('message', async (msg) => {
                const stream = await new Promise((r) => msg.on("body", r));
                const mail = await simpleParser(stream);
                const sender = mail.from?.text || '未知发件人';
                const sendDate = mail.date || new Date().toISOString();
                const folder = CONFIG.TARGET_FOLDERS.chineseName[inboxFolder] || '未知文件夹';
                const verifyCode = getVerifyCodeFromEmail(mail);
                inboxEmail = {
                  sender,
                  sendDate,
                  folder,
                  verifyCode
                };
                res();
              });
            });
          }
        } catch (err) {
          console.error('IMAP获取收件箱邮件失败：', err);
        }

        // 获取垃圾箱邮件及信息
        try {
          const junkFolder = CONFIG.TARGET_FOLDERS.imap[1];
          await new Promise((res, rej) => imap.openBox(junkFolder, true, (err) => err ? rej(err) : res()));
          const junkResults = await new Promise((res, rej) => imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr)));
          if (junkResults.length > 0) {
            const latestJunk = junkResults.slice(-1);
            const f2 = imap.fetch(latestJunk, { bodies: "" });
            await new Promise((res) => {
              f2.on('message', async (msg) => {
                const stream = await new Promise((r) => msg.on("body", r));
                const mail = await simpleParser(stream);
                const sender = mail.from?.text || '未知发件人';
                const sendDate = mail.date || new Date().toISOString();
                const folder = CONFIG.TARGET_FOLDERS.chineseName[junkFolder] || '未知文件夹';
                const verifyCode = getVerifyCodeFromEmail(mail);
                junkEmail = {
                  sender,
                  sendDate,
                  folder,
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

// ===================== 主入口函数（无变更） =====================
module.exports = async (req, res) => {
  try {
    // 方法校验
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).send('不支持的请求方法');
    }

    // 密码校验
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).send('认证失败');
    }

    // 参数校验
    const params = isGet ? req.query : req.body;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      return res.status(400).send(`缺少必要参数：${missingParams.join('、')}`);
    }

    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).send(paramError.message);
    }

    const { refresh_token, client_id, email, response_type = 'json' } = params;

    // 优先使用Graph API
    console.log("【开始】使用Graph API提取验证码及邮件信息");
    const graph_api_result = await graph_api(refresh_token, client_id);
    let emailInfo = null;

    if (graph_api_result.status) {
      emailInfo = await get_dual_folder_latest_email_graph(graph_api_result.access_token);
    } else {
      // 降级使用IMAP
      console.log("【降级】使用IMAP提取验证码及邮件信息");
      const access_token = await get_access_token(refresh_token, client_id);
      const authString = generateAuthString(email, access_token);
      const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };
      emailInfo = await get_dual_folder_latest_email_imap(imapConfig);
    }

    // 无邮件时的响应
    if (!emailInfo) {
      const emptyData = {
        sender: '',
        sendDate: '',
        folder: '',
        verifyCode: ''
      };
      if (response_type === 'html') {
        return res.status(200).send(generateCodeHtmlWithMeta('', '', '', '收件箱/垃圾箱均无邮件'));
      } else {
        return res.status(200).json(emptyData);
      }
    }

    // 有邮件时的响应
    const { sender, sendDate, folder, verifyCode } = emailInfo;
    if (response_type === 'html') {
      return res.status(200).send(generateCodeHtmlWithMeta(verifyCode, sender, sendDate, folder));
    } else {
      return res.status(200).json({
        sender,
        sendDate,
        folder,
        verifyCode
      });
    }

  } catch (error) {
    let statusCode = 500;
    if (error.message.includes('401')) statusCode = 401;
    if (error.message.includes('403')) statusCode = 403;
    if (error.message.includes('请求超时')) statusCode = 504;

    // 异常响应：同步黑色背景风格
    const errorData = {
      sender: '',
      sendDate: '',
      folder: '',
      verifyCode: '',
      error: error.message
    };
    if (req.query.response_type === 'html' || req.body.response_type === 'html') {
      res.status(statusCode).send(`
        <!DOCTYPE html>
        <html lang="zh-CN">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
              body { background: #000; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
              .error-container { color: #ff4444; text-align: center; background: rgba(0,0,0,0.8); padding: 30px; border-radius: 12px; }
              .welcome-title { color: #fff; font-size: 1.8em; font-weight: bold; margin-bottom: 20px; border-bottom: 1px solid #333; padding-bottom: 15px; }
            </style>
          </head>
          <body>
            <div class="error-container">
              <div class="welcome-title">欢迎使用小黑api取件 系统</div>
              <p>错误信息：${error.message}</p>
              <p>验证码：未提取到</p>
            </div>
          </body>
        </html>
      `);
    } else {
      res.status(statusCode).json(errorData);
    }
  }
};
