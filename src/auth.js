const TOKEN_ENDPOINT_TEMPLATE = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

let cachedToken = null;
let refreshPromise = null;

function getRequiredEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error('Missing required environment variable: ' + name);
  }
  return value;
}

async function getAccessToken() {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresAt - 60000 > now) {
    return cachedToken.token;
  }

  if (refreshPromise) return refreshPromise;

  refreshPromise = (async function () {
    try {
      const tenantId = getRequiredEnv('TENANT_ID');
      const clientId = getRequiredEnv('CLIENT_ID');
      const clientSecret = getRequiredEnv('CLIENT_SECRET');

      const params = new URLSearchParams();
      params.set('client_id', clientId);
      params.set('client_secret', clientSecret);
      params.set('grant_type', 'client_credentials');
      params.set('scope', 'https://graph.microsoft.com/.default');

      const response = await fetch(TOKEN_ENDPOINT_TEMPLATE.replace('{tenantId}', tenantId), {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error('Failed to acquire Graph token: ' + text);
      }

      const payload = await response.json();
      cachedToken = {
        token: payload.access_token,
        expiresAt: now + ((payload.expires_in || 3599) * 1000)
      };
      return cachedToken.token;
    } finally {
      refreshPromise = null;
    }
  })();

  return refreshPromise;
}

module.exports = {
  getAccessToken
};
