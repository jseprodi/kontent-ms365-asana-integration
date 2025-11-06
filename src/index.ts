import { observeCustomAppContext, type CustomAppContext } from '@kontent-ai/custom-app-sdk';
import type { AppConfig, SyncContext } from './types/config.js';
import { SyncService } from './services/sync.service.js';
import { getWorkflowStep, getContributors } from './utils/management-api.js';

// Logging helper
function log(level: 'info' | 'warn' | 'error' | 'debug', message: string, data?: any) {
  const prefix = '[Kontent.ai Integration]';
  const logMessage = `${prefix} ${message}`;
  
  if (data) {
    console[level](logMessage, data);
  } else {
    console[level](logMessage);
  }
}

// Parse app configuration from Kontent.ai app config
function getAppConfig(): AppConfig {
  log('info', 'Loading app configuration...');
  
  // In Kontent.ai custom apps, configuration is passed via the appConfig property
  // This will be available when the app is loaded in Kontent.ai
  const appConfigJson = (window as any).__KONTENT_APP_CONFIG__;
  
  let config: AppConfig = {};
  
  if (appConfigJson) {
    try {
      config = typeof appConfigJson === 'string' ? JSON.parse(appConfigJson) : appConfigJson;
      log('info', 'App config loaded from window', { hasConfig: true });
    } catch (error) {
      log('error', 'Failed to parse app config from window:', error);
    }
  } else {
    log('warn', 'No app config found in window.__KONTENT_APP_CONFIG__');
  }

  // Return config with defaults
  const finalConfig: AppConfig = {
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

  log('info', 'Configuration loaded', {
    ms365Enabled: finalConfig.microsoft365?.enabled,
    ms365HasCredentials: !!(finalConfig.microsoft365?.clientId && finalConfig.microsoft365?.tenantId && finalConfig.microsoft365?.clientSecret),
    asanaEnabled: finalConfig.asana?.enabled,
    asanaHasToken: !!finalConfig.asana?.accessToken,
    syncSettings: finalConfig.syncSettings,
  });

  return finalConfig;
}

// Extract sync context from Kontent.ai context
async function extractSyncContext(
  context: CustomAppContext,
  appConfig: AppConfig
): Promise<SyncContext | null> {
  log('debug', 'Extracting sync context', { currentPage: context.currentPage });

  if (context.currentPage !== 'itemEditor') {
    log('debug', 'Not on item editor page, skipping sync');
    return null;
  }

  const syncContext: SyncContext = {
    contentItemId: context.contentItemId,
    languageId: context.languageId,
  };

  log('info', 'Base sync context extracted', {
    contentItemId: syncContext.contentItemId,
    languageId: syncContext.languageId,
  });

  // Fetch additional data from Management API if token is available
  const managementApiToken = (window as any).__KONTENT_MANAGEMENT_API_TOKEN__;
  const projectId = (window as any).__KONTENT_PROJECT_ID__ || context.environmentId;

  log('debug', 'Management API check', {
    hasToken: !!managementApiToken,
    projectId: projectId,
  });

  if (managementApiToken && projectId) {
    try {
      // Fetch workflow step
      if (appConfig.syncSettings?.syncWorkflowSteps) {
        log('info', 'Fetching workflow step from Management API...');
        const workflowStep = await getWorkflowStep(
          context.contentItemId,
          context.languageId,
          managementApiToken,
          projectId
        );
        if (workflowStep) {
          syncContext.workflowStep = workflowStep.codename;
          log('info', 'Workflow step fetched', { workflowStep: workflowStep.codename });
        } else {
          log('warn', 'No workflow step found');
        }
      }

      // Fetch contributors
      if (appConfig.syncSettings?.syncContributors) {
        log('info', 'Fetching contributors from Management API...');
        const contributors = await getContributors(
          context.contentItemId,
          context.languageId,
          managementApiToken,
          projectId
        );
        syncContext.contributors = contributors.map((c) => c.email);
        log('info', 'Contributors fetched', {
          count: contributors.length,
          emails: syncContext.contributors,
        });
      } else {
        log('debug', 'Contributor sync is disabled in settings');
      }
    } catch (error) {
      log('error', 'Failed to fetch additional context from Management API:', error);
    }
  } else {
    log('warn', 'Management API token or project ID not available - cannot fetch workflow steps or contributors', {
      hasToken: !!managementApiToken,
      hasProjectId: !!projectId,
    });
  }

  log('info', 'Final sync context', {
    contentItemId: syncContext.contentItemId,
    languageId: syncContext.languageId,
    workflowStep: syncContext.workflowStep,
    contributors: syncContext.contributors,
    contributorCount: syncContext.contributors?.length || 0,
    dueDate: syncContext.dueDate,
  });

  return syncContext;
}

// Main application entry point
async function initializeApp() {
  log('info', '=== Initializing Kontent.ai Custom App - Microsoft 365 & Asana Integration ===');

  const appConfig = getAppConfig();
  const syncService = new SyncService(appConfig);

  log('info', 'Setting up context observer...');

  // Subscribe to context changes
  const response = await observeCustomAppContext(async (context: CustomAppContext) => {
    log('info', '=== CONTEXT CHANGE DETECTED ===', {
      currentPage: context.currentPage,
      contentItemId: context.currentPage === 'itemEditor' ? context.contentItemId : undefined,
      languageId: context.currentPage === 'itemEditor' ? context.languageId : undefined,
      timestamp: new Date().toISOString(),
    });

    const syncContext = await extractSyncContext(context, appConfig);
    if (syncContext) {
      log('info', 'Starting sync process...', {
        hasContributors: !!(syncContext.contributors && syncContext.contributors.length > 0),
        hasWorkflowStep: !!syncContext.workflowStep,
        hasDueDate: !!syncContext.dueDate,
      });
      
      // Sync to Microsoft 365 and Asana
      syncService.syncContext(syncContext).catch((error) => {
        log('error', 'Failed to sync context:', error);
      });
    } else {
      log('warn', 'No sync context extracted - sync will not occur');
    }
  });

  if (response.isError) {
    log('error', 'Failed to observe context:', {
      errorCode: response.code,
      description: response.description,
    });
    return;
  }

  log('info', 'Context observer initialized successfully');
  log('info', 'Initial context received', {
    currentPage: response.context.currentPage,
    contentItemId: response.context.currentPage === 'itemEditor' ? response.context.contentItemId : undefined,
  });

  // Handle initial context
  const initialSyncContext = await extractSyncContext(response.context, appConfig);
  if (initialSyncContext) {
    log('info', 'Syncing initial context...');
    await syncService.syncContext(initialSyncContext);
  } else {
    log('warn', 'No initial sync context - skipping initial sync');
  }

  // Store unsubscribe function for cleanup
  (window as any).__KONTENT_APP_UNSUBSCRIBE__ = response.unsubscribe;

  log('info', '=== App initialized successfully ===');
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

