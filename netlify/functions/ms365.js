// netlify/functions/ms365.js
import fetch from 'node-fetch';

const REQUIRED_ENV = [
  'MS365_TENANT_ID',
  'MS365_CLIENT_ID',
  'MS365_CLIENT_SECRET',
];

function getMissingEnv() {
  return REQUIRED_ENV.filter((key) => !process.env[key]);
}

async function getAccessToken() {
  const body = new URLSearchParams({
    client_id: process.env.MS365_CLIENT_ID,
    client_secret: process.env.MS365_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${process.env.MS365_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      body,
    }
  );

  if (!tokenRes.ok) {
    const errText = await tokenRes.text();
    throw new Error(`Token request failed: ${tokenRes.status} ${errText}`);
  }

  return tokenRes.json();
}

async function proxyGraphRequest({ userPrincipalName, eventPayload, eventId }) {
  const { access_token } = await getAccessToken();

  const targetUrl = eventId
    ? `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userPrincipalName)}/calendar/events/${eventId}`
    : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userPrincipalName)}/calendar/events`;

  const method = eventId ? 'PATCH' : 'POST';

  const graphRes = await fetch(targetUrl, {
    method,
    headers: {
      'Authorization': `Bearer ${access_token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(eventPayload),
  });

  const text = await graphRes.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    data = text;
  }

  return { status: graphRes.status, ok: graphRes.ok, data };
}

export const handler = async (event) => {
  // Setup CORS (adjust domains as needed)
  const corsHeaders = {
    'Access-Control-Allow-Origin': 'https://kontent-ms365-asana-integration.netlify.app',
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 204,
      headers: corsHeaders,
      body: '',
    };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Method not allowed' }),
    };
  }

  const missingEnv = getMissingEnv();
  if (missingEnv.length) {
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: `Missing env vars: ${missingEnv.join(', ')}` }),
    };
  }

  try {
    const body = JSON.parse(event.body);
    const { userPrincipalName, eventPayload, eventId } = body;

    if (!userPrincipalName || !eventPayload) {
      return {
        statusCode: 400,
        headers: corsHeaders,
        body: JSON.stringify({ error: 'userPrincipalName and eventPayload are required' }),
      };
    }

    const result = await proxyGraphRequest({ userPrincipalName, eventPayload, eventId });

    return {
      statusCode: result.status,
      headers: corsHeaders,
      body: JSON.stringify(result.data),
    };
  } catch (error) {
    console.error('[Netlify Function] Error:', error);
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: error.message }),
    };
  }
};