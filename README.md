# Kontent.ai Custom App - Microsoft 365 & Asana Integration

This custom app integrates Kontent.ai with Microsoft 365 and Asana to automatically create calendar events and tasks when contributor assignments and workflow step updates occur in Kontent.ai.

## Features

- **Microsoft 365 Integration**: Automatically creates calendar events for assigned contributors
- **Asana Integration**: Automatically creates tasks for content items
- **Real-time Sync**: Listens for context changes using the Kontent.ai Custom App SDK
- **Configurable**: Enable/disable integrations and sync settings via configuration

## Prerequisites

- Node.js 18+ and npm
- A Kontent.ai project with Custom Apps enabled
- Microsoft 365 App Registration with Calendar permissions
- Asana Personal Access Token or OAuth token
- (Optional) Kontent.ai Management API token for fetching workflow steps and contributors

## Setup

### 1. Install Dependencies

```powershell
npm install
```

### 2. Configure Environment Variables

Copy `.env.example` to `.env` and fill in your credentials:

```powershell
Copy-Item .env.example .env
```

Edit `.env` with your actual values:

- **Microsoft 365**: Create an Azure AD app registration and grant `Calendars.ReadWrite` permission
- **Asana**: Generate a Personal Access Token from your Asana account settings
- **Kontent.ai Management API**: Optional, but recommended for fetching workflow steps and contributors

### 3. Build the Project

```powershell
npm run build
```

This will compile TypeScript to JavaScript in the `dist` folder.

### 4. Test Locally (Optional)

You can test the app locally using a simple HTTP server:

```powershell
npm run serve
```

Then open `http://localhost:3000` in your browser. Note that the Kontent.ai SDK will only work when loaded within Kontent.ai's iframe.

### 5. Deploy to Kontent.ai

1. Build the project: `npm run build`
2. Upload the `dist` folder contents to your hosting provider (e.g., Azure Static Web Apps, Netlify, Vercel)
3. In Kontent.ai, go to **Project Settings** > **Custom Apps**
4. Click **Add custom app**
5. Configure:
   - **Name**: Microsoft 365 & Asana Integration
   - **URL**: Your hosted app URL
   - **Configuration** (JSON):
   ```json
   {
     "microsoft365": {
       "clientId": "your-client-id",
       "tenantId": "your-tenant-id",
       "clientSecret": "your-client-secret",
       "enabled": true
     },
     "asana": {
       "accessToken": "your-access-token",
       "workspaceId": "your-workspace-id",
       "projectId": "your-project-id",
       "enabled": true
     },
     "syncSettings": {
       "syncContributors": true,
       "syncWorkflowSteps": true,
       "createCalendarEvents": true,
       "createTasks": true
     }
   }
   ```

## Microsoft 365 Setup

**⚠️ Security Note**: The current implementation uses client credentials flow which requires a client secret. Since custom apps run in the browser, exposing secrets is a security risk. For production, consider:

1. **Recommended**: Use a backend API proxy that handles Microsoft 365 authentication
2. **Alternative**: Use delegated permissions with user consent (requires user interaction)

### Option 1: Backend Proxy (Recommended)

Create a backend service that:
- Stores the Microsoft 365 client secret securely
- Handles OAuth token acquisition
- Provides an API endpoint for creating calendar events
- The custom app calls this backend API instead of Microsoft Graph directly

### Option 2: Azure AD App Registration (Current Implementation)

For development/testing only:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: Kontent.ai Integration
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Leave blank (not needed for client credentials flow)
5. After creation, note the **Application (client) ID** and **Directory (tenant) ID**
6. Go to **Certificates & secrets** > **New client secret**
7. Copy the secret value (you won't see it again)
8. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**
9. Add `Calendars.ReadWrite` permission
10. Click **Grant admin consent**

**Warning**: Do not use client secrets in production browser applications. Use a backend proxy instead.

## Asana Setup

1. Go to [Asana Developer Console](https://app.asana.com/0/developer-console)
2. Click **Create new token**
3. Copy the generated token
4. (Optional) Get your Workspace ID and Project ID from Asana URLs or API

## Development

### Run in Development Mode

```powershell
npm run dev
```

This will watch for changes and rebuild automatically.

### Type Checking

```powershell
npm run type-check
```

### Linting

```powershell
npm run lint
```

## How It Works

1. The app uses `observeCustomAppContext()` from the Kontent.ai Custom App SDK to listen for context changes
2. When a content item is opened or updated in the item editor:
   - The app extracts the content item ID, language ID, and other relevant information
   - It syncs this information to Microsoft 365 (creates calendar events)
   - It syncs this information to Asana (creates tasks)
3. The sync is idempotent - if an event/task already exists, it updates it instead of creating a duplicate

## Limitations & Notes

- **Security**: The current Microsoft 365 implementation uses client credentials flow with secrets. For production, use a backend proxy to handle authentication securely.
- The Custom App SDK context doesn't include workflow step or contributor information directly
- To get this data, you'll need to use the Management API (see `src/utils/management-api.ts`)
- Calendar events are created for each contributor email address (must match Azure AD UPN)
- Tasks are created in the specified Asana project (or workspace if no project is specified)
- The app runs in the browser, so all API calls are made client-side
- For Management API access, you'll need to inject the token via a backend service or use a proxy

## Extending the Integration

### Adding Workflow Step Detection

To detect workflow step changes, you'll need to:

1. Store the previous workflow step
2. Compare with the current workflow step
3. Trigger sync when changes are detected

Example:

```typescript
let previousWorkflowStep: string | null = null;

const response = await observeCustomAppContext(async (context) => {
  if (context.currentPage === 'itemEditor') {
    const currentStep = await getWorkflowStep(
      context.contentItemId,
      context.languageId,
      managementApiToken,
      projectId
    );
    
    if (currentStep?.codename !== previousWorkflowStep) {
      // Workflow step changed - trigger sync
      await syncService.syncContext({
        contentItemId: context.contentItemId,
        languageId: context.languageId,
        workflowStep: currentStep?.codename,
      });
      previousWorkflowStep = currentStep?.codename || null;
    }
  }
});
```

### Adding Contributor Detection

Similar to workflow steps, you'll need to fetch contributors from the Management API:

```typescript
const contributors = await getContributors(
  context.contentItemId,
  context.languageId,
  managementApiToken,
  projectId
);

const contributorEmails = contributors.map(c => c.email);
```

## Troubleshooting

### Microsoft 365 Authentication Fails

- Verify your client ID, tenant ID, and client secret are correct
- Ensure you've granted admin consent for the `Calendars.ReadWrite` permission
- Check that the app registration is in the same tenant as the users

### Asana Tasks Not Created

- Verify your access token is valid
- Check that the workspace/project IDs are correct
- Ensure the token has permissions to create tasks in the specified project

### Context Not Updating

- Ensure the app is properly loaded in Kontent.ai
- Check browser console for errors
- Verify the Custom App SDK is properly initialized

## License

MIT

