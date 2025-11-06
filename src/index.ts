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

// Helper function to parse config data
function getAppConfigFromData(config: any): AppConfig {
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
    kontent: {
      managementApiToken: config.kontent?.managementApiToken,
      projectId: config.kontent?.projectId,
    },
    syncSettings: {
      syncContributors: config.syncSettings?.syncContributors ?? true,
      syncWorkflowSteps: config.syncSettings?.syncWorkflowSteps ?? true,
      createCalendarEvents: config.syncSettings?.createCalendarEvents ?? true,
      createTasks: config.syncSettings?.createTasks ?? true,
    },
  };
}

// Parse app configuration from Kontent.ai app config
function getAppConfig(context?: CustomAppContext): AppConfig {
  log('info', 'Loading app configuration...');
  
  let config: AppConfig = {};
  
  // Try to get config from context first (provided by Kontent.ai SDK)
  if (context?.appConfig) {
    try {
      const appConfigData = typeof context.appConfig === 'string' 
        ? JSON.parse(context.appConfig) 
        : context.appConfig;
      config = appConfigData as AppConfig;
      log('info', 'App config loaded from context.appConfig', { hasConfig: true });
    } catch (error) {
      log('error', 'Failed to parse app config from context:', error);
    }
  }
  
  // Fallback to window (for testing/development)
  if (!config.microsoft365?.clientId && !config.asana?.accessToken) {
    const appConfigJson = (window as any).__KONTENT_APP_CONFIG__;
    if (appConfigJson) {
      try {
        const windowConfig = typeof appConfigJson === 'string' 
          ? JSON.parse(appConfigJson) 
          : appConfigJson;
        config = windowConfig as AppConfig;
        log('info', 'App config loaded from window.__KONTENT_APP_CONFIG__', { hasConfig: true });
      } catch (error) {
        log('error', 'Failed to parse app config from window:', error);
      }
    }
    
    // Also check URL parameters as last resort (for manual testing)
    if (!config.microsoft365?.clientId && !config.asana?.accessToken) {
      try {
        const urlParams = new URLSearchParams(window.location.search);
        const configParam = urlParams.get('config');
        if (configParam) {
          const urlConfig = JSON.parse(decodeURIComponent(configParam));
          config = urlConfig as AppConfig;
          log('info', 'App config loaded from URL parameter', { hasConfig: true });
        }
      } catch (error) {
        // Ignore URL parameter parsing errors
      }
    }
    
    if (!config.microsoft365?.clientId && !config.asana?.accessToken) {
      log('warn', 'No app config found in context, window, or URL parameters');
      log('warn', 'Please ensure the configuration JSON is properly saved in Kontent.ai Custom App settings');
    }
  }

  // Return config with defaults - use helper function
  const finalConfig = getAppConfigFromData(config);

  log('info', 'Configuration loaded', {
    ms365Enabled: finalConfig.microsoft365?.enabled,
    ms365HasCredentials: !!(finalConfig.microsoft365?.clientId && finalConfig.microsoft365?.tenantId && finalConfig.microsoft365?.clientSecret),
    asanaEnabled: finalConfig.asana?.enabled,
    asanaHasToken: !!finalConfig.asana?.accessToken,
    hasManagementApiToken: !!finalConfig.kontent?.managementApiToken,
    hasProjectId: !!finalConfig.kontent?.projectId,
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
  // Try from appConfig first, then fallback to window (for development)
  const managementApiToken = appConfig.kontent?.managementApiToken 
    || (window as any).__KONTENT_MANAGEMENT_API_TOKEN__;
  const projectId = appConfig.kontent?.projectId 
    || (window as any).__KONTENT_PROJECT_ID__ 
    || context.environmentId;

  log('debug', 'Management API check', {
    hasToken: !!managementApiToken,
    hasTokenFromConfig: !!appConfig.kontent?.managementApiToken,
    projectId: projectId,
    projectIdFromConfig: appConfig.kontent?.projectId,
    environmentId: context.environmentId,
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

  log('info', 'Setting up context observer...');

  // Store the sync service so we can update it when config changes
  let syncService: SyncService | null = null;
  let currentAppConfig: AppConfig | null = null;
  let currentResponse: { context: CustomAppContext; isError: boolean } | null = null;

  // Subscribe to context changes
  const response = await observeCustomAppContext(async (context: CustomAppContext) => {
    log('info', '=== CONTEXT CHANGE DETECTED ===', {
      currentPage: context.currentPage,
      contentItemId: context.currentPage === 'itemEditor' ? context.contentItemId : undefined,
      languageId: context.currentPage === 'itemEditor' ? context.languageId : undefined,
      hasAppConfig: !!context.appConfig,
      appConfigType: context.appConfig ? typeof context.appConfig : 'undefined',
      appConfigPreview: context.appConfig 
        ? (typeof context.appConfig === 'string' 
            ? context.appConfig.substring(0, 200) 
            : JSON.stringify(context.appConfig).substring(0, 200))
        : undefined,
      contextKeys: Object.keys(context),
      timestamp: new Date().toISOString(),
    });
    
    // Log full context for debugging (only on first context change to avoid spam)
    if (!currentAppConfig) {
      console.log('[Kontent.ai Integration] Full context in callback:', context);
      
      // Check if appConfig exists under a different name
      Object.keys(context).forEach(key => {
        if (key.toLowerCase().includes('config') || key.toLowerCase().includes('app')) {
          console.log(`[Kontent.ai Integration] Found potential config key in callback: ${key}`, (context as any)[key]);
        }
      });
    }

    // Get config from context (it might be available now)
    // Use manually set config if available, otherwise try to get from context
    let appConfig: AppConfig;
    if (currentAppConfig && (currentAppConfig.microsoft365?.clientId || currentAppConfig.asana?.accessToken)) {
      // Use manually set config if it has credentials
      appConfig = currentAppConfig;
      log('info', 'Using manually set configuration');
    } else {
      appConfig = getAppConfig(context);
      
      // Check if appConfig became available in this context update
      if (context.appConfig && !currentAppConfig) {
        log('info', 'App config found in context update!', {
          appConfigType: typeof context.appConfig,
          appConfigPreview: typeof context.appConfig === 'string' 
            ? context.appConfig.substring(0, 200)
            : JSON.stringify(context.appConfig).substring(0, 200),
        });
      }
      
      // Recreate sync service if config changed or if it's the first time
      if (!syncService || JSON.stringify(appConfig) !== JSON.stringify(currentAppConfig)) {
        log('info', 'Creating/updating sync service with new config');
        currentAppConfig = appConfig;
        syncService = new SyncService(appConfig);
      }
    }
    
    // Store current response for manual config function
    currentResponse = { context, isError: false };

    const syncContext = await extractSyncContext(context, appConfig);
    if (syncContext) {
      log('info', 'Starting sync process...', {
        hasContributors: !!(syncContext.contributors && syncContext.contributors.length > 0),
        hasWorkflowStep: !!syncContext.workflowStep,
        hasDueDate: !!syncContext.dueDate,
      });
      
      // Sync to Microsoft 365 and Asana
      if (syncService) {
        syncService.syncContext(syncContext).catch((error) => {
          log('error', 'Failed to sync context:', error);
        });
      } else {
        log('warn', 'Sync service not available, skipping sync');
      }
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
    hasAppConfig: !!response.context.appConfig,
    appConfigType: response.context.appConfig ? typeof response.context.appConfig : 'undefined',
    contextKeys: Object.keys(response.context),
  });
  
  // Log full context separately to see all properties
  console.log('[Kontent.ai Integration] Full context object:', response.context);
  console.log('[Kontent.ai Integration] Context keys:', Object.keys(response.context));
  
  // Check if appConfig exists under a different name
  Object.keys(response.context).forEach(key => {
    if (key.toLowerCase().includes('config') || key.toLowerCase().includes('app')) {
      console.log(`[Kontent.ai Integration] Found potential config key: ${key}`, (response.context as any)[key]);
    }
  });

  // Get config from initial context
  let appConfig = getAppConfig(response.context);
  
  // Log if appConfig property exists but is undefined
  if ('appConfig' in response.context && response.context.appConfig === undefined) {
    log('warn', 'appConfig property exists in context but is undefined - configuration may not be saved in Kontent.ai');
    log('info', 'You can manually set the config by running in console:');
    log('info', 'window.__KONTENT_SET_CONFIG__({your config object here})');
  }
  
  currentAppConfig = appConfig;
  syncService = new SyncService(appConfig);

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
  
  // Store current response
  currentResponse = { context: response.context, isError: false };

  // Expose function to manually set config (for testing/debugging) - defined after response is available
  (window as any).__KONTENT_SET_CONFIG__ = (configJson: string | object) => {
    try {
      const config = typeof configJson === 'string' ? JSON.parse(configJson) : configJson;
      log('info', 'Manually setting app configuration...');
      currentAppConfig = getAppConfigFromData(config);
      syncService = new SyncService(currentAppConfig);
      log('info', 'Configuration manually set, sync service recreated', {
        ms365Enabled: currentAppConfig.microsoft365?.enabled,
        asanaEnabled: currentAppConfig.asana?.enabled,
        hasManagementApiToken: !!currentAppConfig.kontent?.managementApiToken,
      });
      
      // If there's a current context, trigger a sync
      if (syncService && currentResponse && !currentResponse.isError) {
        log('info', 'Triggering sync with manually set configuration...');
        extractSyncContext(currentResponse.context, currentAppConfig).then(syncContext => {
          if (syncContext && syncService) {
            syncService.syncContext(syncContext).catch((error) => {
              log('error', 'Failed to sync after manual config set:', error);
            });
          }
        });
      }
      
      return currentAppConfig;
    } catch (error) {
      log('error', 'Failed to manually set config:', error);
      return null;
    }
  };

  log('info', '=== App initialized successfully ===');
  log('info', 'To manually set configuration, run in console:');
  log('info', 'window.__KONTENT_SET_CONFIG__({your config object})');
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

