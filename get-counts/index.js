const axios = require('axios');
const cookie = require('cookie');

module.exports = async function (context, req) {
    // Cookieからトークンを取得
    const cookies = cookie.parse(req.headers.cookie || '');
    const token = cookies.graph_access_token;

    if (!token) {
        context.res = { status: 401, body: "Unauthorized" };
        return;
    }

    try {
        const fileId = process.env.EXCEL_FILE_ID; // ★ID指定に変更推奨（後述）
        // ファイル名から検索する場合は以下
        // const path = `/me/drive/root:/${process.env.EXCEL_FILENAME}:/workbook/tables/${process.env.TABLE_NAME}/rows`;
        
        // パス指定でのGraph API呼び出し
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${process.env.EXCEL_FILENAME}:/workbook/tables/${process.env.TABLE_NAME}/rows`;
        
        const response = await axios.get(url, {
            headers: { Authorization: `Bearer ${token}` }
        });

        const rows = response.data.value;
        const counts = {};
        
        // 集計ロジック
        rows.forEach(row => {
            const eventName = row.values[0][0];
            if (eventName) counts[eventName] = (counts[eventName] || 0) + 1;
        });

        context.res = {
            headers: { "Content-Type": "application/json" },
            body: counts
        };

    } catch (error) {
        context.log(error);
        context.res = { status: 500, body: error.message };
    }
};