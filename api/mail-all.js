const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// 新增：生成多封邮件的HTML列表页面
function generateEmailsHtml(emailsData) {
  // XSS防护：转义特殊字符
  const escapeHtml = (str) => str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

  // 生成单封邮件的HTML块
  const renderEmailItem = (email, index) => {
    const { send, subject, text, html: emailHtml, date } = email;
    const escapedText = escapeHtml(text || '无内容');
    const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
    const formattedDate = new Date(date).toLocaleString() || '未知日期';

    return `
      <div class="email-item" id="email-${index}">
        <div class="email-header" onclick="toggleEmailContent(${index})">
          <h3 class="email-subject">${escapeHtml(subject || '无主题')}</h3>
          <div class="email-meta">
            <span>发件人：${escapeHtml(send || '未知')}</span>
            <span>日期：${formattedDate}</span>
            <span class="toggle-btn">${index === 0 ? '收起' : '展开'}</span>
          </div>
        </div>
        <div class="email-content" id="content-${index}" style="${index === 0 ? 'display:block' : 'display:none'}">
          ${escapedHtml || `<p>${escapedText}</p>`}
        </div>
      </div>
    `;
  };

  // 拼接所有邮件，生成完整HTML
  const emailsHtml = emailsData.map((email, index) => renderEmailItem(email, index)).join('');

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>邮件列表 - 共${emailsData.length}封</title>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background: #f8f9fa; padding: 20px; }
          .page-title { text-align: center; color: #2d3748; margin-bottom: 30px; font-size: 1.8em; }
          .email-list { max-width: 1000px; margin: 0 auto; gap: 15px; display: flex; flex-direction: column; }
          .email-item { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; }
          .email-header { padding: 15px 20px; background: #f5fafe; cursor: pointer; border-bottom: 1px solid #eee; }
          .email-subject { color: #2d3748; margin-bottom: 8px; font-size: 1.1em; }
          .email-meta { display: flex; gap: 20px; color: #4a5568; font-size: 0.9em; }
          .toggle-btn { margin-left: auto; color: #4299e1; font-weight: 500; }
          .email-content { padding: 20px; color: #1a202c; line-height: 1.8; }
          .email-content p { margin-bottom: 10px; }
          .email-content img { max-width: 100%; height: auto; }
          @media (max-width: 768px) {
            .email-meta { flex-direction: column; gap: 5px; }
            .toggle-btn { margin-left: 0; margin-top: 5px; }
          }
        </style>
        <script>
          // 折叠/展开邮件内容
          function toggleEmailContent(index) {
            const content = document.getElementById(\`content-\${index}\`);
            const btn = document.querySelector(\`#email-\${index} .toggle-btn\`);
            if (content.style.display === 'none') {
              content.style.display = 'block';
              btn.textContent = '收起';
            } else {
              content.style.display = 'none';
              btn.textContent = '展开';
            }
          }
        </script>
      </head>
      <body>
        <h1 class="page-title">邮件列表（共${emailsData.length}封）</h1>
        <div class="email-list">
          ${emailsHtml || '<div style="text-align:center; padding:30px; color:#718096;">未获取到邮件</div>'}
        </div>
      </body>
    </html>
  `;
}

async function get_access_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();

    try {
        const data = JSON.parse(responseText);
        return data.access_token;
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

async function graph_api(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'scope': 'https://graph.microsoft.com/.default'
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();

    try {
        const data = JSON.parse(responseText);

        if (data.scope.indexOf('https://graph.microsoft.com/Mail.ReadWrite') != -1) {
            return {
                access_token: data.access_token,
                status: true
            }
        }

        return {
            access_token: data.access_token,
            status: false
        }
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

async function get_emails(access_token, mailbox) {
    if (!access_token) {
        console.log("Failed to obtain access token'");
        return;
    }

    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=10000`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                "Authorization": `Bearer ${access_token}`
            },
        });

        if (!response.ok) {
            const errorText = await response.text();
            return;
        }

        const responseData = await response.json();
        const emails = responseData.value;

        const response_emails = emails.map(item => ({
            send: item['from']['emailAddress']['address'],
            subject: item['subject'],
            text: item['bodyPreview'],
            html: item['body']['content'],
            date: item['createdDateTime'],
        }));

        return response_emails;

    } catch (error) {
        console.error('Error fetching emails:', error);
        return;
    }
}

module.exports = async (req, res) => {
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({
            error: 'Authentication failed. Please provide valid credentials or contact administrator for access. Refer to API documentation for deployment details.'
        });
    }

    // 新增：获取 response_type 参数（默认 json）
    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    // 检查必要参数
    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    // 验证 response_type 合法性
    if (!['json', 'html'].includes(response_type)) {
        return res.status(400).json({ error: 'Invalid response_type. Use "json" or "html".' });
    }

    try {
        console.log("判断是否graph_api");
        const graph_api_result = await graph_api(refresh_token, client_id);

        if (graph_api_result.status) {
            console.log("是graph_api");

            // 统一邮箱文件夹格式
            if (mailbox != "INBOX" && mailbox != "Junk") mailbox = "inbox";
            if (mailbox === 'INBOX') mailbox = 'inbox';
            if (mailbox === 'Junk') mailbox = 'junkemail';

            const result = await get_emails(graph_api_result.access_token, mailbox);

            // 支持 HTML/JSON 响应
            if (response_type === 'html') {
                const htmlResponse = generateEmailsHtml(result || []);
                res.status(200).send(htmlResponse);
            } else {
                res.status(200).json(result);
            }
            return;
        }

        // IMAP 流程处理
        const access_token = await get_access_token(refresh_token, client_id);
        const authString = generateAuthString(email, access_token);
        const emailList = [];

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: {
                rejectUnauthorized: false
            }
        });

        imap.once("ready", async () => {
            try {
                await new Promise((resolve, reject) => {
                    imap.openBox(mailbox, true, (err, box) => {
                        if (err) return reject(err);
                        resolve(box);
                    });
                });

                const results = await new Promise((resolve, reject) => {
                    imap.search(["ALL"], (err, results) => {
                        if (err) return reject(err);
                        resolve(results);
                    });
                });

                const f = imap.fetch(results, { bodies: "" });

                f.on("message", (msg, seqno) => {
                    msg.on("body", (stream, info) => {
                        simpleParser(stream, (err, mail) => {
                            if (err) throw err;
                            emailList.push({
                                send: mail.from.text,
                                subject: mail.subject,
                                text: mail.text,
                                html: mail.html,
                                date: mail.date,
                            });
                        });
                    });
                });

                f.once("end", () => {
                    imap.end();
                });
            } catch (err) {
                imap.end();
                res.status(500).json({ error: err.message });
            }
        });

        imap.once('error', (err) => {
            console.error('IMAP error:', err);
            res.status(500).json({ error: err.message });
        });

        imap.once('end', () => {
            console.log('IMAP connection ended');
            // IMAP 流程支持 HTML/JSON 响应
            if (response_type === 'html') {
                const htmlResponse = generateEmailsHtml(emailList);
                res.status(200).send(htmlResponse);
            } else {
                res.status(200).json(emailList);
            }
        });

        imap.connect();

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
};
