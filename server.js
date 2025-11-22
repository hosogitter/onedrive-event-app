// ▼ この行を追加してください（SSL証明書エラーを無視する設定）
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

require('dotenv').config();
const express = require('express');

require('dotenv').config();

const session = require('express-session');
const msal = require('@azure/msal-node');
const axios = require('axios');

const app = express();
app.use(express.static('public'));
app.use(express.json());
app.use(session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false
}));

// --- 1. MSAL (認証) 設定 ---
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

// ログイン用URL生成とリダイレクト
app.get('/auth/signin', async (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read", "files.readwrite"],
        redirectUri: process.env.REDIRECT_URI,
    };
    const url = await cca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(url);
});

// コールバック処理 (トークン取得)
app.get('/auth/callback', async (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read", "files.readwrite"],
        redirectUri: process.env.REDIRECT_URI,
    };

    try {
        const response = await cca.acquireTokenByCode(tokenRequest);
        req.session.accessToken = response.accessToken; // セッションに保存
        res.redirect('/'); // トップページへ戻る
    } catch (error) {
        res.status(500).send(error);
    }
});

// --- 2. API処理 (フロントエンドから呼ばれる) ---

// Graph APIを呼ぶためのヘルパー関数 (ファイルパスから操作)
const getGraphClient = (token) => axios.create({
    baseURL: 'https://graph.microsoft.com/v1.0',
    headers: { Authorization: `Bearer ${token}` }
});

// [API] データのカウントを取得
app.get('/api/counts', async (req, res) => {
    if (!req.session.accessToken) return res.status(401).json({ error: 'Unauthorized' });

    try {
        const client = getGraphClient(req.session.accessToken);
        // ファイル名からパス指定でテーブル行を取得
        const path = `/me/drive/root:/${process.env.EXCEL_FILENAME}:/workbook/tables/${process.env.TABLE_NAME}/rows`;
        
        const response = await client.get(path);
        const rows = response.data.value;

        // イベントごとの集計ロジック (0列目がイベント名と仮定)
        const counts = {};
        rows.forEach(row => {
            const eventName = row.values[0][0];
            if (eventName) counts[eventName] = (counts[eventName] || 0) + 1;
        });

        res.json(counts);
    } catch (error) {
        console.error(error.response?.data || error.message);
        res.status(500).json({ error: 'Error fetching data' });
    }
});

// [API] データの追加
app.post('/api/add', async (req, res) => {
    if (!req.session.accessToken) return res.status(401).json({ error: 'Unauthorized' });

    try {
        const client = getGraphClient(req.session.accessToken);
        const path = `/me/drive/root:/${process.env.EXCEL_FILENAME}:/workbook/tables/${process.env.TABLE_NAME}/rows`;

        const { eventName, name, tel, receptionist, date } = req.body;

        // 追加データ (Excelの列順に合わせる: イベント, 名前, Tel, 受付者, 日付)
        const newRow = [[eventName, name, tel, receptionist, date]];

        await client.post(path, { values: newRow });
        res.json({ success: true });

    } catch (error) {
        console.error(error.response?.data || error.message);
        res.status(500).json({ error: 'Error adding data' });
    }
});

app.listen(3000, () => console.log('Server running on http://localhost:3000'));