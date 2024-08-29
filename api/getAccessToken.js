const { google } = require('google-auth-library');

// OAuth2 credentials
const clientId = '';
const clientSecret = '';
const redirectUri = 'YOUR_REDIRECT_URI';

// Create an OAuth2 client
const oauth2Client = new google.auth.OAuth2(
  clientId,
  clientSecret,
  redirectUri
);

// Generate the URL for authorization
const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline', // For refresh token
  scope: ['https://www.googleapis.com/auth/script.external_request'],
});

console.log('Authorize this app by visiting this url:', authUrl);

// Assuming you get the code from the user after they visit the authUrl
const getAccessToken = async (code) => {
  const { tokens } = await oauth2Client.getToken(code);
  oauth2Client.setCredentials(tokens);
  console.log('Access Token:', tokens.access_token);
  return tokens.access_token;
};

// Example usage
// Replace 'YOUR_AUTHORIZATION_CODE' with the actual authorization code
getAccessToken('p')
  .then(accessToken => {
    // Use the access token to make a request to your Web App
    const webAppUrl = 'YOUR_WEB_APP_URL';

    const fetch = require('node-fetch');
    fetch(webAppUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    })
    .then(res => res.text())
    .then(body => console.log('Response from Web App:', body))
    .catch(error => console.error('Error:', error));
  })
  .catch(error => console.error('Error retrieving access token:', error));
