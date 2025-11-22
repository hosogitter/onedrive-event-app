const axios = require('axios');
const cookie = require('cookie');

module.exports = async function (context, req) {
    const cookies = cookie.parse(req.headers.cookie || '');
    const token = cookies.graph_access_token;

    if (!token) {
        context.res = { status: 401, body: "Unauthorized" };
        return;
    }

    const { eventName, name, tel, receptionist, date } = req.body;

    try {
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${process.env.EXCEL_FILENAME}:/workbook/tables/${process.env.TABLE_NAME}/rows`;

        const newRow = [[eventName, name, tel, receptionist, date]];

        await axios.post(url, { values: newRow }, {
            headers: { Authorization: `Bearer ${token}` }
        });

        context.res = { body: { success: true } };

    } catch (error) {
        context.log(error);
        context.res = { status: 500, body: error.message };
    }
};