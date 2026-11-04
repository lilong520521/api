// ============================================================================
// /api/mail-all —— 获取全部邮件（Graph 优先，IMAP 兜底）
// ----------------------------------------------------------------------------
// 本次修复的 6 个 Vercel 致命问题（旧版直接 500 / 返回空数组）：
//   FIX-1 依赖地狱：旧版 require('node-imap')，但 package.json 里没有这个包。
//          Vercel 安装依赖时不会装它 → 函数一加载就抛 Cannot find module
//          → FUNCTION_INVOCATION_FAILED（"获取全部邮件失败"）。
//          现在统一改用 package.json 里已存在的 imapflow，零新增依赖。
//   FIX-2 mailparser 3.x 取消了回调：旧版 simpleParser(stream, cb) 的回调永远不触发
//          → emailList 恒为空 → 返回 []。现在全部 await。
//   FIX-3 Graph 分页上限：旧版 $top=10000，Graph 单页最大 1000，超过直接 400
//          → 取件函数静默 return undefined → 响应体是 undefined。现在按 1000 内分页。
//   FIX-4 响应体超限：Vercel Serverless 响应上限 4.5MB，一次拉 10000 封 HTML 必然炸。
//          现在每封 HTML 截断 120KB + 总量预算 3.5MB。
//   FIX-5 邮箱文件夹映射写反：旧版 `if (mailbox != "INBOX" && mailbox != "Junk") mailbox = "inbox"`
//          把 junkemail / 收件箱 / 已发送 全部强改成 inbox。现在中英别名统一映射 + IMAP 特殊文件夹探测。
//   FIX-6 双重响应：旧版 imap 'error' 和 'end' 都会 res.json → Cannot set headers after they are sent。
//          现在统一收口到 sendJson / sendText，只发一次。
// ============================================================================

const { ImapFlow } = require('imapflow');
const simpleParser = require('mailparser').simpleParser;

const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_MAIL_FOLDERS: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  REQUEST_TIMEOUT: 20000,        // 单个请求超时（ms）
  GRAPH_TIMEOUT: 25000,          // 拉列表整轮超时预算（ms）
  DEFAULT_LIMIT: 50,             // 默认返回封数
  MAX_LIMIT: 200,                // 允许的最大封数
  GRAPH_PAGE_SIZE: 50,           // 每页请求封数（必须 <= 1000）
  MAX_HTML_PER_MAIL: 120 * 1024, // 单封 HTML 正文截断长度
  MAX_RESPONSE_BYTES: 3.5 * 1024 * 1024, // 总响应预算（Vercel 上限 4.5MB）
  IMAP: {
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    connectionTimeout: 15000,
    greetingTimeout: 10000,
    socketTimeout: 30000
  }
};

// ---------------------------------------------------------------- 工具函数
async function fetchWithTimeout(url, options = {}, timeout = CONFIG.REQUEST_TIMEOUT) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (err) {
    const msg = err && err.name === 'AbortError' ? `请求超时（${timeout}ms）` : (err && err.message) || String(err);
    throw new Error(msg);
  } finally {
    clearTimeout(timer);
  }
}

function escapeHtml(str) {
  if (str === undefined || str === null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function clampInt(value, fallback, min, max) {
  const n = parseInt(value, 10);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

// 把用户传来的 mailbox 统一成规范 key：inbox / junkemail / sentitems / drafts / deleteditems / archive
const FOLDER_ALIASES = {
  inbox: 'inbox', in: 'inbox', '收件箱': 'inbox', '收信': 'inbox',
  junk: 'junkemail', junkemail: 'junkemail', junkemailfolder: 'junkemail', spam: 'junkemail',
  '垃圾箱': 'junkemail', '垃圾邮件': 'junkemail', '垃圾邮箱': 'junkemail',
  sent: 'sentitems', sentitems: 'sentitems', 'sent items': 'sentitems', '已发送': 'sentitems', '已发送邮件': 'sentitems',
  draft: 'drafts', drafts: 'drafts', '草稿': 'drafts',
  deleteditems: 'deleteditems', deleted: 'deleteditems', trash: 'deleteditems',
  '已删除': 'deleteditems', '已删除邮件': 'deleteditems', '删除邮件': 'deleteditems',
  archive: 'archive', '存档': 'archive', '归档': 'archive'
};

function normalizeFolder(input) {
  const raw = String(input || 'inbox').trim().toLowerCase();
  return FOLDER_ALIASES[raw] || 'inbox';
}

const FOLDER_CN = {
  inbox: '收件箱',
  junkemail: '垃圾箱',
  sentitems: '已发送',
  drafts: '草稿',
  deleteditems: '已删除',
  archive: '存档'
};

// IMAP 特殊文件夹：先用 SPECIAL-USE 标志探测，探测不到再用英文名兜底
const IMAP_SPECIAL_USE = {
  inbox: '\\Inbox',
  junkemail: '\\Junk',
  sentitems: '\\Sent',
  drafts: '\\Drafts',
  deleteditems: '\\Trash',
  archive: '\\Archive'
};

const IMAP_FALLBACK_PATH = {
  inbox: 'INBOX',
  junkemail: 'Junk',
  sentitems: 'Sent Items',
  drafts: 'Drafts',
  deleteditems: 'Deleted Items',
  archive: 'Archive'
};

// ---------------------------------------------------------------- 令牌
async function getAccessToken(refresh_token, client_id, scope) {
  const body = new URLSearchParams({
    client_id: client_id,
    grant_type: 'refresh_token',
    refresh_token: refresh_token
  });
  if (scope) body.set('scope', scope);

  const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  const text = await response.text();
  let data = null;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error(`令牌接口返回非 JSON：${text.slice(0, 300)}`);
  }

  if (!response.ok) {
    const errCode = data && data.error ? data.error : `HTTP ${response.status}`;
    const errDesc = data && data.error_description ? data.error_description : text.slice(0, 300);
    throw new Error(`令牌获取失败：${errCode} ${String(errDesc).split('\n')[0].slice(0, 300)}`);
  }

  return data;
}

// Graph 权限探活：Mail.Read / Mail.ReadWrite 任一即可；scope 为空也放行（与 mail-new 对齐）
function hasGraphMailScope(scopeStr) {
  const s = String(scopeStr || '');
  if (!s) return true;
  return s.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1 ||
    s.indexOf('https://graph.microsoft.com/Mail.Read') !== -1 ||
    /(^|\s)Mail\.Read(Write)?(\s|$)/i.test(s);
}

// ---------------------------------------------------------------- Graph 取件
function mapGraphMessage(item, folderKey) {
  const htmlRaw = (item.body && item.body.content) || '';
  const text = item.bodyPreview || '';
  // Outlook 有时会给纯文本正文，发 HTML 反而更好读
  let html = htmlRaw;
  if (html.length > CONFIG.MAX_HTML_PER_MAIL) {
    html = html.slice(0, CONFIG.MAX_HTML_PER_MAIL) +
      '<p style="color:#888">…（正文过长，已截断，可在邮箱客户端查看完整内容）</p>';
  }
  if (!html) {
    html = `<p>${escapeHtml(text).replace(/\n/g, '<br>')}</p>`;
  }

  return {
    send: (item.from && item.from.emailAddress && (item.from.emailAddress.address || item.from.emailAddress.name)) || '未知发件人',
    subject: item.subject || '无主题',
    text: text,
    html: html,
    date: item.receivedDateTime || item.createdDateTime || new Date().toISOString(),
    folderSource: folderKey
  };
}

async function graphFetchMessages(accessToken, folderKey, limit) {
  const select = 'from,subject,bodyPreview,body,receivedDateTime,createdDateTime';
  const pageSize = Math.min(limit, CONFIG.GRAPH_PAGE_SIZE);
  let url = `${CONFIG.GRAPH_MAIL_FOLDERS}/${encodeURIComponent(folderKey)}/messages` +
    `?$top=${pageSize}&$orderby=${encodeURIComponent('receivedDateTime desc')}&$select=${select}`;

  const collected = [];
  const deadline = Date.now() + CONFIG.GRAPH_TIMEOUT;
  let pages = 0;

  while (url && collected.length < limit && pages < 10) {
    if (Date.now() > deadline) break;

    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Prefer': 'outlook.body-content-type="html"'
      }
    }, CONFIG.REQUEST_TIMEOUT);

    if (!response.ok) {
      const errText = await response.text();
      const hint = response.status === 401 ? '令牌无 Mail 权限或已过期'
        : response.status === 403 ? '应用缺少 Mail.ReadWrite/Mail.Read 权限（需在 Azure 后台授予并同意）'
        : response.status === 404 ? `文件夹 ${folderKey} 不存在`
        : 'Graph 返回异常';
      throw new Error(`Graph ${response.status} ${hint}：${errText.slice(0, 300)}`);
    }

    const json = await response.json();
    const values = Array.isArray(json.value) ? json.value : [];
    for (const item of values) {
      if (collected.length >= limit) break;
      collected.push(mapGraphMessage(item, FOLDER_CN[folderKey] || folderKey));
    }

    url = json['@odata.nextLink'] || null;
    pages += 1;
  }

  return collected;
}

// ---------------------------------------------------------------- IMAP 取件
async function resolveImapPath(client, folderKey) {
  let boxes = [];
  try {
    boxes = await client.list();
  } catch (e) {
    boxes = [];
  }

  const wantSpecial = (IMAP_SPECIAL_USE[folderKey] || '').toLowerCase();
  if (wantSpecial) {
    const hit = boxes.find(b => String(b.specialUse || '').toLowerCase() === wantSpecial);
    if (hit) return hit.path;
  }

  const fallback = IMAP_FALLBACK_PATH[folderKey] || 'INBOX';
  const byPath = boxes.find(b => String(b.path || '').toLowerCase() === fallback.toLowerCase());
  return byPath ? byPath.path : fallback;
}

// imapflow 的 "Command failed" 只是外壳，真实原因在 responseStatus / responseText 里
function describeImapError(err) {
  if (!err) return 'unknown';
  const parts = [];
  parts.push(err.message || String(err));
  if (err.responseStatus) parts.push(`command=${err.responseStatus}`);
  if (err.serverResponseCode) parts.push(`serverCode=${err.serverResponseCode}`);
  if (err.responseText) parts.push(`serverText=${String(err.responseText).slice(0, 300)}`);
  if (err.authenticationFailed) parts.push('authenticationFailed=true');
  if (err.code) parts.push(`code=${err.code}`);
  return parts.join(' | ');
}

async function imapFetchMessages(email, accessToken, folderKey, limit) {
  const client = new ImapFlow({
    host: CONFIG.IMAP.host,
    port: CONFIG.IMAP.port,
    secure: CONFIG.IMAP.secure,
    auth: { user: email, accessToken: accessToken },
    logger: false,
    tls: { rejectUnauthorized: false },
    connectionTimeout: CONFIG.IMAP.connectionTimeout,
    greetingTimeout: CONFIG.IMAP.greetingTimeout,
    socketTimeout: CONFIG.IMAP.socketTimeout
  });

  // imapflow 在 logger:false 时偶发抛出 "Logging is disabled" 的噪音错误，吞掉它
  if (typeof client.emit === 'function' && typeof client.on === 'function') {
    const origEmit = client.emit.bind(client);
    client.emit = function (event, ...args) {
      if (event === 'error' && args[0]) {
        const m = args[0].message || String(args[0]);
        if (m.includes('Logging is disabled')) return false;
      }
      return origEmit(event, ...args);
    };
  }

  const out = [];
  try {
    await client.connect();
    const path = await resolveImapPath(client, folderKey);

    const lock = await client.getMailboxLock(path);
    try {
      const total = client.mailbox.exists || 0;
      if (total > 0) {
        const seqStart = Math.max(1, total - limit + 1);
        for await (const msg of client.fetch(`${seqStart}:${total}`, { envelope: true, source: true })) {
          let parsed = null;
          try {
            parsed = await simpleParser(msg.source);
          } catch (e) {
            parsed = null;
          }

          const env = msg.envelope || {};
          const fromAddr = (parsed && parsed.from && parsed.from.text)
            || (env.from && env.from[0] && (env.from[0].address || env.from[0].name))
            || '未知发件人';
          const subject = (parsed && parsed.subject) || env.subject || '无主题';
          const text = (parsed && parsed.text) || '';
          let html = (parsed && parsed.html) || '';
          if (typeof html === 'string' && html.length > CONFIG.MAX_HTML_PER_MAIL) {
            html = html.slice(0, CONFIG.MAX_HTML_PER_MAIL) +
              '<p style="color:#888">…（正文过长，已截断，可在邮箱客户端查看完整内容）</p>';
          }
          if (!html) html = `<p>${escapeHtml(text).replace(/\n/g, '<br>')}</p>`;

          const when = (parsed && parsed.date) || (env.date) || new Date();
          out.push({
            send: fromAddr,
            subject: subject,
            text: text,
            html: html,
            date: new Date(when).toISOString(),
            folderSource: FOLDER_CN[folderKey] || folderKey
          });
        }
      }
    } finally {
      lock.release();
    }

    try { await client.logout(); } catch (e) { /* 已断开 */ }
  } catch (err) {
    try { await client.logout(); } catch (e) { /* 忽略 */ }
    throw err;
  }

  // IMAP 顺序是旧的在前，倒过来让最新的排第一（与 Graph 行为一致）
  return out.reverse();
}

// ---------------------------------------------------------------- 响应体积预算
function enforceByteBudget(list) {
  const kept = [];
  let used = 0;
  for (const item of list) {
    const size = Buffer.byteLength(JSON.stringify(item), 'utf8');
    if (used + size > CONFIG.MAX_RESPONSE_BYTES) {
      // 超预算：先丢掉正文 HTML 只留 text，尽量多保留几封
      const slim = { ...item, html: `<p>${escapeHtml(item.text || '').slice(0, 2000).replace(/\n/g, '<br>')}</p>` };
      const slimSize = Buffer.byteLength(JSON.stringify(slim), 'utf8');
      if (used + slimSize > CONFIG.MAX_RESPONSE_BYTES) break;
      kept.push(slim);
      used += slimSize;
      continue;
    }
    kept.push(item);
    used += size;
  }
  return kept;
}

// ---------------------------------------------------------------- HTML 列表页
function generateEmailsHtml(emailsData) {
  const list = Array.isArray(emailsData) ? emailsData : [];

  const renderEmailItem = (email, index) => `
      <div class="email-item" id="email-${index}">
        <div class="email-header" onclick="toggleEmailContent(${index})">
          <h3 class="email-subject">${escapeHtml(email.subject || '无主题')}</h3>
          <div class="email-meta">
            <span>发件人：${escapeHtml(email.send || '未知')}</span>
            <span>日期：${escapeHtml(new Date(email.date).toLocaleString())}</span>
            <span class="toggle-btn">${index === 0 ? '收起' : '展开'}</span>
          </div>
        </div>
        <div class="email-content" id="content-${index}" style="${index === 0 ? 'display:block' : 'display:none'}">
          ${email.html || `<p>${escapeHtml(email.text || '无内容').replace(/\n/g, '<br>')}</p>`}
        </div>
      </div>`;

  return `<!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>邮件列表 - 共${list.length}封</title>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background: #f8f9fa; padding: 20px; }
          .page-title { text-align: center; color: #2d3748; margin-bottom: 30px; font-size: 1.8em; }
          .email-list { max-width: 1000px; margin: 0 auto; gap: 15px; display: flex; flex-direction: column; }
          .email-item { background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,.08); overflow: hidden; }
          .email-header { padding: 15px 20px; background: #f5fafe; cursor: pointer; border-bottom: 1px solid #eee; }
          .email-subject { color: #2d3748; margin-bottom: 8px; font-size: 1.1em; }
          .email-meta { display: flex; gap: 20px; color: #4a5568; font-size: .9em; }
          .toggle-btn { margin-left: auto; color: #4299e1; font-weight: 500; }
          .email-content { padding: 20px; color: #1a202c; line-height: 1.8; overflow-wrap: anywhere; }
          .email-content img { max-width: 100%; height: auto; }
          @media (max-width: 768px) {
            .email-meta { flex-direction: column; gap: 5px; }
            .toggle-btn { margin-left: 0; margin-top: 5px; }
          }
        </style>
        <script>
          function toggleEmailContent(index) {
            var content = document.getElementById('content-' + index);
            var btn = document.querySelector('#email-' + index + ' .toggle-btn');
            if (!content) return;
            if (content.style.display === 'none') { content.style.display = 'block'; if (btn) btn.textContent = '收起'; }
            else { content.style.display = 'none'; if (btn) btn.textContent = '展开'; }
          }
        </script>
      </head>
      <body>
        <h1 class="page-title">邮件列表（共${list.length}封）</h1>
        <div class="email-list">
          ${list.map(renderEmailItem).join('') || '<div style="text-align:center;padding:30px;color:#718096;">未获取到邮件</div>'}
        </div>
      </body>
    </html>`;
}

// ---------------------------------------------------------------- 主入口
module.exports = async (req, res) => {
  let step = 'init';
  let sent = false;
  const sendJson = (status, payload) => {
    if (sent) return;
    sent = true;
    res.status(status).json(payload);
  };
  const sendText = (status, body) => {
    if (sent) return;
    sent = true;
    res.status(status).send(body);
  };

  try {
    step = '1.方法校验';
    if (!['GET', 'POST'].includes(req.method)) {
      return sendJson(405, { error: 'Method Not Allowed，请使用 GET 或 POST' });
    }

    step = '2.参数解析';
    let params = req.method === 'GET' ? (req.query || {}) : (req.body || {});
    if (typeof params === 'string') {
      try { params = JSON.parse(params); } catch (e) { return sendJson(400, { error: '请求体不是合法 JSON' }); }
    }

    step = '3.密码校验';
    const expectedPassword = process.env.PASSWORD;
    if (expectedPassword && params.password !== expectedPassword) {
      return sendJson(401, { error: '密码错误。请在环境变量 PASSWORD 中核对，或联系管理员获取密码。' });
    }

    step = '4.缺少参数';
    const refresh_token = params.refresh_token;
    const client_id = params.client_id;
    const email = params.email;
    const missing = ['refresh_token', 'client_id', 'email'].filter(k => !params[k]);
    if (missing.length) {
      return sendJson(400, { error: `缺少必要参数：${missing.join('、')}` });
    }

    const response_type = String(params.response_type || 'json').toLowerCase();
    if (!['json', 'html'].includes(response_type)) {
      return sendJson(400, { error: 'response_type 只能是 json 或 html' });
    }

    const folderKey = normalizeFolder(params.mailbox);
    const limit = clampInt(params.limit, CONFIG.DEFAULT_LIMIT, 1, CONFIG.MAX_LIMIT);

    // ---- Graph 优先 ----
    step = '5.Graph令牌';
    let messages = [];
    let graphError = '';
    try {
      const tokenData = await getAccessToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
      if (hasGraphMailScope(tokenData.scope)) {
        step = '6.Graph取件';
        messages = await graphFetchMessages(tokenData.access_token, folderKey, limit);
      } else {
        graphError = `Graph 令牌未包含 Mail 权限，scope=${tokenData.scope}`;
      }
    } catch (err) {
      graphError = `Graph 通道失败：${err.message}`;
    }

    // ---- Graph 拿不到就兜底 IMAP（与 mail-new 的降级逻辑一致）----
    let imapError = '';
    if (messages.length === 0) {
      step = '7.IMAP兜底';
      try {
        const tokenData = await getAccessToken(refresh_token, client_id);
        messages = await imapFetchMessages(email, tokenData.access_token, folderKey, limit);
      } catch (err) {
        imapError = `IMAP 通道失败：${describeImapError(err)}`;
      }
    }

    // ---- 两条路都报错 → 500，把两边的真实原因都吐出来，方便你在 Vercel 日志里定位 ----
    if (messages.length === 0 && graphError && imapError) {
      return sendJson(500, {
        error: `未能获取邮件。${graphError} ｜ ${imapError}`,
        graphError: graphError,
        imapError: imapError,
        folder: folderKey,
        limit: limit
      });
    }

    step = '8.体积预算';
    const finalList = enforceByteBudget(messages);

    step = '9.响应';
    if (response_type === 'html') {
      return sendText(200, generateEmailsHtml(finalList));
    }
    // 前端（public/mail.html、index.html）要求这里是纯数组
    return sendJson(200, finalList);

  } catch (error) {
    const msg = (error && error.message) || String(error);
    console.error(`[mail-all] 步骤[${step}]失败：${msg}`);
    return sendJson(500, { error: `步骤[${step}]失败：${msg}` });
  }
};
