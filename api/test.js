const { google } = require('googleapis');

async function callApiExecutable() {
    const oauth2Client = new google.auth.OAuth2(
        'YOUR_CLIENT',
        'YOUR_SECRET',
        'YOUR_REDIRECT_URL'
    );

    // 获取访问令牌
    oauth2Client.setCredentials({ refresh_token: 'YOUR_REFRESH_TOKEN' });

    const scriptId = '18Os5bP8YSpMWjxQS8HDB8Q5neD7A9gk_IIBvMyxFFFj11344D_nPZmSe'; // 替换为您的脚本 ID
    const request = {
        resource: {
            function: 'doGet',
            parameters: [{ name: 'Alice' }]
        },
        auth: oauth2Client,
    };

    const script = google.script({ version: 'v1', auth: oauth2Client });
    const response = await script.scripts.run({
        scriptId: scriptId,
        resource: request.resource,
    });

    console.log('Response:', response.data);
}

callApiExecutable().catch(console.error);