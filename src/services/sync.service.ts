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
    const syncKey = `${context.contentItemId}-${context.languageId}`;
    const existingSync = this.syncMap.get(syncKey);

    const syncSettings = this.config.syncSettings || {
      syncContributors: true,
      syncWorkflowSteps: true,
      createCalendarEvents: true,
      createTasks: true,
    };

    // Sync to Microsoft 365
    if (syncSettings.createCalendarEvents && this.ms365Service.isEnabled()) {
      await this.syncToMicrosoft365(context, existingSync?.ms365EventId);
    }

    // Sync to Asana
    if (syncSettings.createTasks && this.asanaService.isEnabled()) {
      await this.syncToAsana(context, existingSync?.asanaTaskId);
    }

    // Update sync map
    const updatedSync = {
      ms365EventId: existingSync?.ms365EventId,
      asanaTaskId: existingSync?.asanaTaskId,
    };
    this.syncMap.set(syncKey, updatedSync);
  }

  private async syncToMicrosoft365(context: SyncContext, existingEventId?: string): Promise<void> {
    if (!this.config.microsoft365?.enabled) return;

    // Calculate event times (default to 1 hour event, or use due date if available)
    const startTime = context.dueDate
      ? new Date(context.dueDate.getTime() - 30 * 60 * 1000) // 30 minutes before due date
      : new Date(Date.now() + 24 * 60 * 60 * 1000); // Tomorrow if no due date
    const endTime = context.dueDate
      ? new Date(context.dueDate.getTime() + 30 * 60 * 1000) // 30 minutes after due date
      : new Date(startTime.getTime() + 60 * 60 * 1000); // 1 hour duration

    // Sync for each contributor if available
    // Note: For Microsoft 365, we need user principal names (UPN) or user IDs, not just emails
    // The email should be the UPN format: user@domain.com
    const contributors = context.contributors || [];
    if (contributors.length === 0) {
      console.warn('No contributors specified, skipping Microsoft 365 sync');
      return;
    }

    for (const contributorEmail of contributors) {
      try {
        // Use email as user principal name (UPN) - this should match the user's UPN in Azure AD
        const userPrincipalName = contributorEmail;
        
        if (existingEventId) {
          await this.ms365Service.updateCalendarEvent(
            userPrincipalName,
            existingEventId,
            context,
            startTime,
            endTime
          );
        } else {
          const eventId = await this.ms365Service.createCalendarEvent(
            userPrincipalName,
            context,
            startTime,
            endTime
          );
          if (eventId) {
            const syncKey = `${context.contentItemId}-${context.languageId}`;
            const existing = this.syncMap.get(syncKey) || {};
            existing.ms365EventId = eventId;
            this.syncMap.set(syncKey, existing);
          }
        }
      } catch (error) {
        console.error(`Failed to sync to Microsoft 365 for ${contributorEmail}:`, error);
      }
    }
  }

  private async syncToAsana(context: SyncContext, existingTaskId?: string): Promise<void> {
    if (!this.config.asana?.enabled) return;

    try {
      // Use first contributor as assignee if available
      const assigneeEmail = context.contributors?.[0];

      if (existingTaskId) {
        await this.asanaService.updateTask(existingTaskId, context, assigneeEmail);
      } else {
        const taskId = await this.asanaService.createTask(context, assigneeEmail);
        if (taskId) {
          const syncKey = `${context.contentItemId}-${context.languageId}`;
          const existing = this.syncMap.get(syncKey) || {};
          existing.asanaTaskId = taskId;
          this.syncMap.set(syncKey, existing);
        }
      }
    } catch (error) {
      console.error('Failed to sync to Asana:', error);
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

