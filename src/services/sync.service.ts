import type { AppConfig, SyncContext } from '../types/config.js';
import { Microsoft365Service } from './microsoft365.service.js';
import { AsanaService } from './asana.service.js';

export class SyncService {
  private ms365Service: Microsoft365Service;
  private asanaService: AsanaService;
  private config: AppConfig;
  private syncMap: Map<string, { ms365EventId?: string; asanaTaskId?: string }> = new Map();

  constructor(config: AppConfig) {
    this.config = config;
    this.ms365Service = new Microsoft365Service(config.microsoft365);
    this.asanaService = new AsanaService(config.asana);
  }

  async syncContext(context: SyncContext): Promise<void> {
    console.log('[SyncService] Starting sync', {
      contentItemId: context.contentItemId,
      languageId: context.languageId,
      contributors: context.contributors,
      workflowStep: context.workflowStep,
      dueDate: context.dueDate,
    });

    const syncKey = `${context.contentItemId}-${context.languageId}`;
    const existingSync = this.syncMap.get(syncKey);

    console.log('[SyncService] Existing sync status', {
      syncKey,
      hasExistingSync: !!existingSync,
      existingMs365EventId: existingSync?.ms365EventId,
      existingAsanaTaskId: existingSync?.asanaTaskId,
    });

    const syncSettings = this.config.syncSettings || {
      syncContributors: true,
      syncWorkflowSteps: true,
      createCalendarEvents: true,
      createTasks: true,
    };

    console.log('[SyncService] Sync settings', syncSettings);
    console.log('[SyncService] Service status', {
      ms365Enabled: this.ms365Service.isEnabled(),
      asanaEnabled: this.asanaService.isEnabled(),
    });

    // Sync to Microsoft 365
    if (syncSettings.createCalendarEvents && this.ms365Service.isEnabled()) {
      console.log('[SyncService] Syncing to Microsoft 365...');
      await this.syncToMicrosoft365(context, existingSync?.ms365EventId);
    } else {
      console.log('[SyncService] Skipping Microsoft 365 sync', {
        createCalendarEventsEnabled: syncSettings.createCalendarEvents,
        ms365ServiceEnabled: this.ms365Service.isEnabled(),
      });
    }

    // Sync to Asana
    if (syncSettings.createTasks && this.asanaService.isEnabled()) {
      console.log('[SyncService] Syncing to Asana...');
      await this.syncToAsana(context, existingSync?.asanaTaskId);
    } else {
      console.log('[SyncService] Skipping Asana sync', {
        createTasksEnabled: syncSettings.createTasks,
        asanaServiceEnabled: this.asanaService.isEnabled(),
      });
    }

    // Update sync map
    const updatedSync = {
      ms365EventId: existingSync?.ms365EventId,
      asanaTaskId: existingSync?.asanaTaskId,
    };
    this.syncMap.set(syncKey, updatedSync);

    console.log('[SyncService] Sync completed', {
      syncKey,
      updatedSync,
    });
  }

  private async syncToMicrosoft365(context: SyncContext, existingEventId?: string): Promise<void> {
    console.log('[SyncService] syncToMicrosoft365 called', {
      enabled: this.config.microsoft365?.enabled,
      existingEventId,
      hasDueDate: !!context.dueDate,
      contributors: context.contributors,
    });

    if (!this.config.microsoft365?.enabled) {
      console.log('[SyncService] Microsoft 365 is disabled, skipping');
      return;
    }

    // Calculate event times (default to 1 hour event, or use due date if available)
    const startTime = context.dueDate
      ? new Date(context.dueDate.getTime() - 30 * 60 * 1000) // 30 minutes before due date
      : new Date(Date.now() + 24 * 60 * 60 * 1000); // Tomorrow if no due date
    const endTime = context.dueDate
      ? new Date(context.dueDate.getTime() + 30 * 60 * 1000) // 30 minutes after due date
      : new Date(startTime.getTime() + 60 * 60 * 1000); // 1 hour duration

    console.log('[SyncService] Calculated event times', {
      startTime: startTime.toISOString(),
      endTime: endTime.toISOString(),
      dueDate: context.dueDate?.toISOString(),
    });

    // Sync for each contributor if available
    // Note: For Microsoft 365, we need user principal names (UPN) or user IDs, not just emails
    // The email should be the UPN format: user@domain.com
    const contributors = context.contributors || [];
    console.log('[SyncService] Contributors to sync', { count: contributors.length, emails: contributors });

    if (contributors.length === 0) {
      console.warn('[SyncService] No contributors specified, skipping Microsoft 365 sync');
      return;
    }

    for (const contributorEmail of contributors) {
      try {
        console.log('[SyncService] Processing contributor', { email: contributorEmail });
        
        // Use email as user principal name (UPN) - this should match the user's UPN in Azure AD
        const userPrincipalName = contributorEmail;
        
        if (existingEventId) {
          console.log('[SyncService] Updating existing calendar event', {
            userPrincipalName,
            eventId: existingEventId,
          });
          await this.ms365Service.updateCalendarEvent(
            userPrincipalName,
            existingEventId,
            context,
            startTime,
            endTime
          );
          console.log('[SyncService] Calendar event updated successfully');
        } else {
          console.log('[SyncService] Creating new calendar event', { userPrincipalName });
          const eventId = await this.ms365Service.createCalendarEvent(
            userPrincipalName,
            context,
            startTime,
            endTime
          );
          if (eventId) {
            console.log('[SyncService] Calendar event created', { eventId });
            const syncKey = `${context.contentItemId}-${context.languageId}`;
            const existing = this.syncMap.get(syncKey) || {};
            existing.ms365EventId = eventId;
            this.syncMap.set(syncKey, existing);
          } else {
            console.warn('[SyncService] Calendar event creation returned no event ID');
          }
        }
      } catch (error) {
        console.error(`[SyncService] Failed to sync to Microsoft 365 for ${contributorEmail}:`, error);
      }
    }
  }

  private async syncToAsana(context: SyncContext, existingTaskId?: string): Promise<void> {
    console.log('[SyncService] syncToAsana called', {
      enabled: this.config.asana?.enabled,
      existingTaskId,
      hasContributors: !!(context.contributors && context.contributors.length > 0),
    });

    if (!this.config.asana?.enabled) {
      console.log('[SyncService] Asana is disabled, skipping');
      return;
    }

    try {
      // Use first contributor as assignee if available
      const assigneeEmail = context.contributors?.[0];
      console.log('[SyncService] Asana assignee', { assigneeEmail });

      if (existingTaskId) {
        console.log('[SyncService] Updating existing Asana task', { taskId: existingTaskId });
        const success = await this.asanaService.updateTask(existingTaskId, context, assigneeEmail);
        console.log('[SyncService] Asana task update result', { success });
      } else {
        console.log('[SyncService] Creating new Asana task');
        const taskId = await this.asanaService.createTask(context, assigneeEmail);
        if (taskId) {
          console.log('[SyncService] Asana task created', { taskId });
          const syncKey = `${context.contentItemId}-${context.languageId}`;
          const existing = this.syncMap.get(syncKey) || {};
          existing.asanaTaskId = taskId;
          this.syncMap.set(syncKey, existing);
        } else {
          console.warn('[SyncService] Asana task creation returned no task ID');
        }
      }
    } catch (error) {
      console.error('[SyncService] Failed to sync to Asana:', error);
    }
  }

  getSyncStatus(contentItemId: string, languageId: string): {
    ms365EventId?: string;
    asanaTaskId?: string;
  } | null {
    const syncKey = `${contentItemId}-${languageId}`;
    return this.syncMap.get(syncKey) || null;
  }
}

