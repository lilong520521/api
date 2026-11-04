// ============================================================================
// /api/imap-probe —— IMAP 通道诊断（只读，不写任何东西）
// 目的：当 /api/mail-all 报 "IMAP 通道失败: Command failed" 时，把真正的原因挖出来。
//   imapflow 的 "Command failed" 只是外壳，真实信息在 err.responseStatus / err.responseText
//   / err.executedCommand 里，普通代码只打了 err.message 就丢掉了。
//
// 用法（把四个参数换成你的）：
//   GET /api/imap-probe?refresh_token=xxx&client_id=yyy&email=a@outlook.com
//   GET /api/imap-probe?refresh_token=xxx&client_id=yyy&email=a@outlook.com&password=你的PASSWORD
//
// 返回三段 attempts：
//   1) imapflow（默认客户端标识）        → 和你线上取件用的完全一样
//   2) imapflow（clientInfo: {} 简化 ID）→ 看是不是 ID 命令被拒
//   3) raw-tls（裸 TLS + XOAUTH2 内联）  → 绕开 imapflow，验证服务器是否根本不接受这台机器的会话
// 三段里哪一段成功，就说明该用哪条路。
// ============================================================================

const { ImapFlow } = require('imapflow');
const tls = require('tls');

const TOKEN_URLS = [
  'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  'https://login.microsoftonline.com/common/oauth2/v2.0/token',
];

function describeError(err) {
  if (!err) return { message: 'unknown' };
  return {
    message: err.message,
    code: err.code,
    responseStatus: err.responseStatus,        // 失败的是哪个 IMAP 命令（ID / NAMESPACE / SELECT / AUTHENTICATE…）
    responseText: err.responseText,            // 服务器原话
    serverResponseCode: err.serverResponseCode,
    executedCommand: err.executedCommand,      // 实际发出去的完整命令行
    authenticationFailed: !!err.authenticationFailed,
    stack: String(err.stack || '').split('\n').slice(0, 4).join(' | '),
  };
}

async function getImapToken(refresh_token, client_id) {
  let last = null;
  for (const url of TOKEN_URLS) {
    const body = new URLSearchParams({
      client_id,
      grant_type: 'refresh_token',
      refresh_token,
    });
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    const text = await resp.text();
    let data = null;
    try { data = JSON.parse(text); } catch (e) { data = null; }
    if (data && data.access_token) return { ok: true, token: data.access_token, scope: data.scope || '', endpoint: url };
    last = { ok: false, status: resp.status, body: text.slice(0, 300) };
    if (data && data.error && !/7000012|different tenant/i.test(String(data.error_description || ''))) break;
  }
  return last;
}

async function attemptImapflow(email, token, extra = {}, label = 'imapflow') {
  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: { user: email, accessToken: token },
    logger: false,
    tls: { rejectUnauthorized: false },
    connectionTimeout: 20000,
    greetingTimeout: 15000,
    socketTimeout: 30000,
    ...extra,
  });

  const origEmit = client.emit.bind(client);
  client.emit = function (event, ...args) {
    if (event === 'error' && args[0]) {
      const m = args[0].message || String(args[0]);
      if (m.includes('Logging is disabled')) return false;
    }
    return origEmit(event, ...args);
  };

  const out = { label, ok: false };
  try {
    await client.connect();
    out.connected = true;
    const boxes = await client.list();
    out.mailboxes = boxes.length;
    out.mailboxNames = boxes.slice(0, 8).map(b => b.path + (b.specialUse ? '[' + b.specialUse + ']' : ''));
    const lock = await client.getMailboxLock('INBOX');
    try {
      out.inboxCount = client.mailbox.exists || 0;
    } finally {
      lock.release();
    }
    out.ok = true;
  } catch (err) {
    out.error = describeError(err);
  } finally {
    try { await client.logout(); } catch (e) { }
  }
  return out;
}

// 裸 TLS + IMAP：CAPABILITY → AUTHENTICATE XOAUTH2（内联）→ SELECT INBOX → UID SEARCH ALL
function attemptRawTls(email, token) {
  return new Promise((resolve) => {
    const out = { label: 'raw-tls', ok: false, log: [] };
    let buf = '';
    let done = false;
    let sock = null;

    const finish = (ok, err) => {
      if (done) return;
      done = true;
      out.ok = ok;
      if (err) out.error = err;
      try { if (sock) sock.destroy(); } catch (e) { }
      resolve(out);
    };

    const timer = setTimeout(() => finish(false, { message: '整体超时（45s）' }), 45000);

    try {
      sock = tls.connect({ host: 'outlook.office365.com', port: 993, rejectUnauthorized: false, servername: 'outlook.office365.com' }, () => { });
    } catch (e) {
      clearTimeout(timer);
      return finish(false, { message: 'TLS 连接失败: ' + e.message });
    }

    sock.setTimeout(30000);
    sock.on('timeout', () => { clearTimeout(timer); finish(false, { message: 'socket 超时' }); });
    sock.on('error', (e) => { clearTimeout(timer); finish(false, { message: 'socket error: ' + e.message }); });

    const waitTag = (tag) => new Promise((res) => {
      const check = () => {
        const re = new RegExp('(^|\\r\\n)' + tag + ' (OK|NO|BAD)[^\\r\\n]*');
        const m = buf.match(re);
        if (m) { res(m[0].trim()); return true; }
        return false;
      };
      if (check()) return;
      const iv = setInterval(() => { if (check()) clearInterval(iv); }, 50);
      const guard = setTimeout(() => { clearInterval(iv); res('(等待 ' + tag + ' 超时，收到: ' + buf.slice(-160).replace(/\r?\n/g, ' / ') + ')'); }, 20000);
      sock.once('close', () => { setTimeout(() => { clearInterval(iv); clearTimeout(guard); res('(连接被关闭，收到: ' + buf.slice(-160).replace(/\r?\n/g, ' / ') + ')'); }, 100); });
    });

    sock.on('data', (d) => { buf += d.toString('utf8'); });

    (async () => {
      try {
        // 等 greeting
        await new Promise((res) => {
          const iv = setInterval(() => { if (buf.includes('\r\n')) { clearInterval(iv); res(); } }, 50);
          setTimeout(() => { clearInterval(iv); res(); }, 8000);
        });
        out.log.push('greeting: ' + buf.split('\r\n')[0].slice(0, 120));

        sock.write('p1 CAPABILITY\r\n');
        out.log.push('CAPABILITY: ' + String(await waitTag('p1')).slice(0, 160));

        const b64 = Buffer.from(`user=${email}\x01auth=Bearer ${token}\x01\x01`).toString('base64');
        sock.write(`p2 AUTHENTICATE XOAUTH2 ${b64}\r\n`);
        const authLine = await waitTag('p2');
        out.log.push('AUTHENTICATE: ' + String(authLine).slice(0, 200));
        if (!/^p2 OK/i.test(String(authLine))) {
          clearTimeout(timer);
          return finish(false, { message: 'AUTHENTICATE 被拒', responseText: String(authLine).slice(0, 300) });
        }

        sock.write('p3 SELECT "INBOX"\r\n');
        const sel = await waitTag('p3');
        out.log.push('SELECT: ' + String(sel).slice(0, 160));
        if (!/^p3 OK/i.test(String(sel))) {
          clearTimeout(timer);
          return finish(false, { message: 'SELECT 被拒', responseText: String(sel).slice(0, 300) });
        }

        sock.write('p4 UID SEARCH ALL\r\n');
        const se = await waitTag('p4');
        out.log.push('SEARCH: ' + String(se).slice(0, 120));
        const m = buf.match(/\* SEARCH([0-9 ]*)/);
        out.uidCount = m ? m[1].trim().split(/\s+/).filter(Boolean).length : 0;
        clearTimeout(timer);
        finish(true);
      } catch (e) {
        clearTimeout(timer);
        finish(false, { message: '流程异常: ' + e.message });
      }
    })();
  });
}

module.exports = async (req, res) => {
  const params = req.method === 'GET' ? (req.query || {}) : (req.body || {});
  const { refresh_token, client_id, email, password } = params;

  if (!refresh_token || !client_id || !email) {
    return res.status(400).json({ error: '缺少必要参数：refresh_token、client_id、email' });
  }
  if (process.env.PASSWORD && password !== process.env.PASSWORD) {
    return res.status(401).json({ error: '密码错误' });
  }

  const started = Date.now();
  const result = {
    email,
    region: process.env.VERCEL_REGION || '(unknown)',
    node: process.version,
    imapflowVersion: (() => { try { return require('imapflow/package.json').version; } catch (e) { return '?'; } })(),
  };

  const tk = await getImapToken(refresh_token, client_id);
  if (!tk.ok) {
    result.token = tk;
    return res.status(200).json(result);
  }
  result.token = { ok: true, endpoint: tk.endpoint, scope: (tk.scope || '').slice(0, 200) };

  result.attempts = [];
  result.attempts.push(await attemptImapflow(email, tk.token, {}, 'imapflow-default'));
  if (!result.attempts[0].ok) {
    result.attempts.push(await attemptImapflow(email, tk.token, { clientInfo: {} }, 'imapflow-empty-id'));
  }
  if (!result.attempts.some(a => a.ok)) {
    result.attempts.push(await attemptRawTls(email, tk.token));
  }

  result.elapsedMs = Date.now() - started;
  result.verdict = result.attempts.some(a => a.ok)
    ? '至少有一条通道可用：' + result.attempts.filter(a => a.ok).map(a => a.label).join(' / ')
    : '三条通道全部失败，看 attempts 里的 responseStatus / responseText';
  return res.status(200).json(result);
};
