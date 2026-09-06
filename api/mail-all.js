const { ImapFlow } = require('imapflow');
const simpleParser = require("mailparser").simpleParser;

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

        if (data.scope.indexOf('https://graph.microsoft.com/Mail.Read') != -1) {
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
            return
        }

        const responseData = await response.json();

        const emails = responseData.value;

        const response_emails = emails.map(item => {
            return {
                send: item['from']['emailAddress']['address'],
                subject: item['subject'],
                text: item['bodyPreview'],
                html: item['body']['content'],
                date: item['createdDateTime'],
            }
        })

        return response_emails

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
            error: '密码验证失败'
        });
    }

    let { refresh_token, client_id, email, mailbox } = req.method === 'GET' ? req.query : req.body;

    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    const ALLOWED_MAILBOXES = ['INBOX', 'Junk'];
    if (!ALLOWED_MAILBOXES.includes(mailbox)) {
        return res.status(400).json({ error: 'Invalid mailbox. Allowed: INBOX, Junk' });
    }

    try {

        let useGraph = false;
        let graphAccessToken = null;
        try {
            const r = await graph_api(refresh_token, client_id);
            if (r && r.status) { useGraph = true; graphAccessToken = r.access_token; }
        } catch (e) {
            console.log("Graph API 未授权，将使用 IMAP:", e.message.split('\n')[0]);
        }

        if (useGraph && graphAccessToken) {
            console.log("使用 Graph API");
            if (mailbox == 'INBOX') mailbox = 'inbox';
            if (mailbox == 'Junk') mailbox = 'junkemail';
            const result = await get_emails(graphAccessToken, mailbox);
            res.status(200).json(result || []);
            return
        }

        // ===== IMAP 路径：imapflow 替代 node-imap =====
        const access_token = await get_access_token(refresh_token, client_id);
        const client = new ImapFlow({
            host: 'outlook.office365.com', port: 993, secure: true,
            auth: { user: email, accessToken: access_token },
            logger: false, greetingTimeout: 10000, socketTimeout: 60000
        });
        await client.connect();
        const emailList = [];
        const lock = await client.getMailboxLock(mailbox);
        try {
            const exists = client.mailbox && client.mailbox.exists;
            if (exists > 0) {
                for await (const msg of client.fetch({ all: true }, { source: true })) {
                    try {
                        const mail = await simpleParser(msg.source);
                        emailList.push({
                            send: mail.from && mail.from.text,
                            subject: mail.subject,
                            text: mail.text,
                            html: mail.html,
                            date: mail.date,
                        });
                    } catch (parseErr) {
                        console.error('message parse error:', parseErr.message);
                    }
                }
            }
        } finally { lock.release(); }
        try { await client.logout(); } catch (e) { client.close(); }
        console.log('IMAP done, count:', emailList.length);
        res.status(200).json(emailList);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
};
