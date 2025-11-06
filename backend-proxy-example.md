# Backend Proxy Example for Microsoft 365 Authentication

This document provides an example of how to create a secure backend proxy for Microsoft 365 authentication, since client secrets should not be exposed in browser applications.

## Why a Backend Proxy?

- **Security**: Client secrets should never be exposed in browser code
- **Token Management**: Backend can handle token refresh and caching
- **Rate Limiting**: Backend can implement rate limiting and error handling
- **CORS**: Backend can handle CORS properly for your custom app

## Example Implementation (Node.js/Express)

```typescript
import express from 'express';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

const app = express();
app.use(express.json());

// Store these in environment variables, not in code
const MS365_CLIENT_ID = process.env.MS365_CLIENT_ID!;
const MS365_TENANT_ID = process.env.MS365_TENANT_ID!;
const MS365_CLIENT_SECRET = process.env.MS365_CLIENT_SECRET!;

// Cache for access tokens
let cachedToken: { token: string; expiresAt: number } | null = null;

async function getAccessToken(): Promise<string> {
  // Return cached token if still valid
  if (cachedToken && cachedToken.expiresAt > Date.now()) {
    return cachedToken.token;
  }

  // Get new token
  const tokenEndpoint = `https://login.microsoftonline.com/${MS365_TENANT_ID}/oauth2/v2.0/token`;
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      client_id: MS365_CLIENT_ID,
      client_secret: MS365_CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });

  if (!response.ok) {
    throw new Error(`Failed to get access token: ${response.statusText}`);
  }

  const data = await response.json();
  const expiresIn = data.expires_in * 1000; // Convert to milliseconds
  cachedToken = {
    token: data.access_token,
    expiresAt: Date.now() + expiresIn,
  };

  return cachedToken.token;
}

// Initialize Graph client
async function getGraphClient(): Promise<Client> {
  const token = await getAccessToken();
  return Client.init({
    authProvider: (done) => {
      done(null, token);
    },
  });
}

// CORS middleware
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*'); // In production, specify your domain
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// Create calendar event endpoint
app.post('/api/calendar/events', async (req, res) => {
  try {
    const { userId, subject, startTime, endTime, body } = req.body;

    if (!userId || !subject || !startTime || !endTime) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const client = await getGraphClient();
    const event = {
      subject,
      body: {
        contentType: 'HTML',
        content: body || '',
      },
      start: {
        dateTime: startTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: endTime,
        timeZone: 'UTC',
      },
      isReminderOn: true,
      reminderMinutesBeforeStart: 15,
    };

    const createdEvent = await client
      .api(`/users/${userId}/calendar/events`)
      .post(event);

    res.json({ id: createdEvent.id, ...createdEvent });
  } catch (error: any) {
    console.error('Error creating calendar event:', error);
    res.status(500).json({ error: error.message });
  }
});

// Update calendar event endpoint
app.patch('/api/calendar/events/:eventId', async (req, res) => {
  try {
    const { eventId } = req.params;
    const { userId, subject, startTime, endTime, body } = req.body;

    if (!userId) {
      return res.status(400).json({ error: 'Missing userId' });
    }

    const client = await getGraphClient();
    const event: any = {};

    if (subject) event.subject = subject;
    if (body) {
      event.body = {
        contentType: 'HTML',
        content: body,
      };
    }
    if (startTime) {
      event.start = {
        dateTime: startTime,
        timeZone: 'UTC',
      };
    }
    if (endTime) {
      event.end = {
        dateTime: endTime,
        timeZone: 'UTC',
      };
    }

    await client.api(`/users/${userId}/calendar/events/${eventId}`).patch(event);

    res.json({ success: true });
  } catch (error: any) {
    console.error('Error updating calendar event:', error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Backend proxy server running on port ${PORT}`);
});
```

## Updating the Custom App to Use the Proxy

Update `src/services/microsoft365.service.ts` to call your backend proxy instead:

```typescript
async createCalendarEvent(
  userId: string,
  context: SyncContext,
  startTime: Date,
  endTime: Date
): Promise<string | null> {
  if (!this.config?.enabled) {
    return null;
  }

  try {
    const proxyUrl = process.env.BACKEND_PROXY_URL || 'http://localhost:3001';
    const response = await fetch(`${proxyUrl}/api/calendar/events`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        userId,
        subject: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
        startTime: startTime.toISOString(),
        endTime: endTime.toISOString(),
        body: `
          <p>Content Item ID: ${context.contentItemId}</p>
          <p>Language ID: ${context.languageId}</p>
          ${context.workflowStep ? `<p>Workflow Step: ${context.workflowStep}</p>` : ''}
          ${context.contributors ? `<p>Contributors: ${context.contributors.join(', ')}</p>` : ''}
        `,
      }),
    });

    if (!response.ok) {
      throw new Error(`Failed to create calendar event: ${response.statusText}`);
    }

    const data = await response.json();
    return data.id;
  } catch (error) {
    console.error('Failed to create calendar event:', error);
    return null;
  }
}
```

## Deployment

1. Deploy the backend proxy to a secure server (Azure App Service, AWS Lambda, etc.)
2. Set environment variables securely (Azure Key Vault, AWS Secrets Manager, etc.)
3. Update your custom app configuration with the proxy URL
4. Remove Microsoft 365 client secret from the custom app configuration

## Security Best Practices

1. **Use HTTPS**: Always use HTTPS for the proxy endpoint
2. **Authentication**: Add API key or OAuth authentication to your proxy
3. **Rate Limiting**: Implement rate limiting to prevent abuse
4. **Input Validation**: Validate all inputs on the backend
5. **Error Handling**: Don't expose sensitive error messages to clients
6. **CORS**: Restrict CORS to your custom app domain only

