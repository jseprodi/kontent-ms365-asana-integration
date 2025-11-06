import { observeCustomAppContext, type CustomAppContext } from '@kontent-ai/custom-app-sdk';
import type { AppConfig, SyncContext } from './types/config.js';
import { SyncService } from './services/sync.service.js';
import { getWorkflowStep, getContributors } from './utils/management-api.js';

// Parse app configuration from Kontent.ai app config
function getAppConfig(): AppConfig {
  // In Kontent.ai custom apps, configuration is passed via the appConfig property
  // This will be available when the app is loaded in Kontent.ai
  const appConfigJson = (window as any).__KONTENT_APP_CONFIG__;
  
  let config: AppConfig = {};
  
  if (appConfigJson) {
    try {
      config = typeof appConfigJson === 'string' ? JSON.parse(appConfigJson) : appConfigJson;
    } catch (error) {
      console.error('Failed to parse app config from window:', error);
    }
  }

  // Return config with defaults
  return {
    microsoft365: {
      clientId: config.microsoft365?.clientId || '',
      tenantId: config.microsoft365?.tenantId || '',
      clientSecret: config.microsoft365?.clientSecret || '',
      enabled: config.microsoft365?.enabled ?? false,
    },
    asana: {
      accessToken: config.asana?.accessToken || '',
      workspaceId: config.asana?.workspaceId,
      projectId: config.asana?.projectId,
      enabled: config.asana?.enabled ?? false,
    },
    syncSettings: {
      syncContributors: config.syncSettings?.syncContributors ?? true,
      syncWorkflowSteps: config.syncSettings?.syncWorkflowSteps ?? true,
      createCalendarEvents: config.syncSettings?.createCalendarEvents ?? true,
      createTasks: config.syncSettings?.createTasks ?? true,
    },
  };
}

// Extract sync context from Kontent.ai context
async function extractSyncContext(
  context: CustomAppContext,
  appConfig: AppConfig
): Promise<SyncContext | null> {
  if (context.currentPage !== 'itemEditor') {
    return null;
  }

  const syncContext: SyncContext = {
    contentItemId: context.contentItemId,
    languageId: context.languageId,
  };

  // Fetch additional data from Management API if token is available
  const managementApiToken = (window as any).__KONTENT_MANAGEMENT_API_TOKEN__;
  const projectId = (window as any).__KONTENT_PROJECT_ID__ || context.environmentId;

  if (managementApiToken && projectId) {
    try {
      // Fetch workflow step
      if (appConfig.syncSettings?.syncWorkflowSteps) {
        const workflowStep = await getWorkflowStep(
          context.contentItemId,
          context.languageId,
          managementApiToken,
          projectId
        );
        if (workflowStep) {
          syncContext.workflowStep = workflowStep.codename;
        }
      }

      // Fetch contributors
      if (appConfig.syncSettings?.syncContributors) {
        const contributors = await getContributors(
          context.contentItemId,
          context.languageId,
          managementApiToken,
          projectId
        );
        syncContext.contributors = contributors.map((c) => c.email);
      }
    } catch (error) {
      console.warn('Failed to fetch additional context from Management API:', error);
    }
  }

  return syncContext;
}

// Main application entry point
async function initializeApp() {
  console.log('Initializing Kontent.ai Custom App - Microsoft 365 & Asana Integration');

  const appConfig = getAppConfig();
  const syncService = new SyncService(appConfig);

  // Subscribe to context changes
  const response = await observeCustomAppContext(async (context: CustomAppContext) => {
    console.log('Context updated:', context);

    const syncContext = await extractSyncContext(context, appConfig);
    if (syncContext) {
      // Sync to Microsoft 365 and Asana
      syncService.syncContext(syncContext).catch((error) => {
        console.error('Failed to sync context:', error);
      });
    }
  });

  if (response.isError) {
    console.error('Failed to observe context:', {
      errorCode: response.code,
      description: response.description,
    });
    return;
  }

  console.log('Initial context:', response.context);

  // Handle initial context
  const initialSyncContext = await extractSyncContext(response.context, appConfig);
  if (initialSyncContext) {
    await syncService.syncContext(initialSyncContext);
  }

  // Store unsubscribe function for cleanup
  (window as any).__KONTENT_APP_UNSUBSCRIBE__ = response.unsubscribe;

  console.log('App initialized successfully');
}

// Initialize app when DOM is ready
if (typeof window !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
  } else {
    initializeApp();
  }
}

// Export for potential external use
export { initializeApp, getAppConfig };

