const msal = require('@azure/msal-node');

// 設定は環境変数から読み込む
const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};
const cca = new msal.ConfidentialClientApplication(config);

module.exports = async function (context, req) {
    const authCodeUrlParameters = {
        scopes: ["user.read", "files.readwrite"],
        // Azure SWAのドメインに合わせて後で調整が必要
        redirectUri: process.env.REDIRECT_URI, 
    };

    try {
        const url = await cca.getAuthCodeUrl(authCodeUrlParameters);
        // ユーザーをMicrosoftのログイン画面へ飛ばす
        context.res = {
            status: 302,
            headers: { "Location": url },
            body: {}
        };
    } catch (error) {
        context.res = { status: 500, body: "Auth Init Error" };
    }
};