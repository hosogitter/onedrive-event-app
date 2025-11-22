const msal = require('@azure/msal-node');
const cookie = require('cookie');

const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};
const cca = new msal.ConfidentialClientApplication(config);

module.exports = async function (context, req) {
    const code = req.query.code;
    
    if (!code) {
        context.res = { status: 400, body: "No code provided" };
        return;
    }

    const tokenRequest = {
        code: code,
        scopes: ["user.read", "files.readwrite"],
        redirectUri: process.env.REDIRECT_URI,
    };

    try {
        const response = await cca.acquireTokenByCode(tokenRequest);
        const accessToken = response.accessToken;

        // トークンをCookieとしてブラウザに保存させる設定
        const cookieStr = cookie.serialize('graph_access_token', accessToken, {
            httpOnly: true,
            secure: true, // HTTPS必須
            maxAge: 3600, // 1時間
            path: '/'
        });

        // トップページへ戻す
        context.res = {
            status: 302,
            headers: {
                "Location": "/", 
                "Set-Cookie": cookieStr
            },
            body: {}
        };
    } catch (error) {
        context.log(error);
        context.res = { status: 500, body: "Token Error" };
    }
};