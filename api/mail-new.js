const { ImapFlow } = require('imapflow');

// ========================================================================
// 【核心替换】用 imapflow 替代 node-imap
// ========================================================================
async function get_dual_folder_latest_email_imap(imapConfig) {
  // imapConfig 里带的是 { user, xoauth2 }，imapflow 需要 accessToken
  // 所以这个函数接收原始参数
  const { email, accessToken } = imapConfig;

  const result = {
    email: null,
    error: '',
    diagnostics: { inboxCount: 0, junkCount: 0 }
  };

  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: {
      user: email,
      accessToken: accessToken
    },
    logger: false,                     // 👈 关键：完全禁用日志
    tls: { rejectUnauthorized: false },
    connectionTimeout: 15000,
    greetingTimeout: 10000,
    socketTimeout: 20000
  });

  // 屏蔽 imapflow 内部可能残留的日志
  try {
    const origEmit = client.emit.bind(client);
    client.emit = function(event, ...args) {
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

  let inboxEmail = null;
  let junkEmail = null;

  try {
    // 连接
    await client.connect();

    // ---- 收件箱 ----
    try {
      const lock = await client.getMailboxLock('INBOX');
      try {
        const total = client.mailbox.exists || 0;
        result.diagnostics.inboxCount = total;
        if (total > 0) {
          // 取最后一封（seq 号从 1 开始，最新的是 total）
          for await (const msg of client.fetch({ seq: String(total) }, { envelope: true, source: true })) {
            const mail = await simpleParser(msg.source);
            const verifyCode = getVerifyCodeFromEmail(
              { text: mail.text, html: mail.html }, mail.subject
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
          for await (const msg of client.fetch({ seq: String(total) }, { envelope: true, source: true })) {
            const mail = await simpleParser(msg.source);
            const verifyCode = getVerifyCodeFromEmail(
              { text: mail.text, html: mail.html }, mail.subject
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

    // 关闭
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
    try { await client.logout(); } catch (e) {}
    return result;
  }
}
