import * as asana from 'asana';
import type { AppConfig, SyncContext } from '../types/config.js';

export class AsanaService {
  private client: asana.Client | null = null;
  private config: AppConfig['asana'];
  private projectId: string | null = null;

  constructor(config: AppConfig['asana']) {
    this.config = config;
    if (config?.enabled && config.accessToken) {
      this.initializeClient();
    }
  }

  private initializeClient() {
    if (!this.config?.accessToken) return;

    try {
      this.client = asana.Client.create().useAccessToken(this.config.accessToken);

      // Use project ID from config if provided
      if (this.config.projectId) {
        this.projectId = this.config.projectId;
      }
    } catch (error) {
      console.error('Failed to initialize Asana client:', error);
      throw error;
    }
  }

  async createTask(context: SyncContext, assigneeEmail?: string): Promise<string | null> {
    if (!this.client || !this.config?.enabled) {
      console.warn('Asana service is not enabled or not initialized');
      return null;
    }

    try {
      const taskData: any = {
        name: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
        notes: `
Content Item ID: ${context.contentItemId}
Language ID: ${context.languageId}
${context.workflowStep ? `Workflow Step: ${context.workflowStep}\n` : ''}
${context.contributors ? `Contributors: ${context.contributors.join(', ')}\n` : ''}
        `.trim(),
        due_on: context.dueDate?.toISOString().split('T')[0], // Asana expects YYYY-MM-DD format
      };

      // Add project if configured
      if (this.projectId) {
        taskData.projects = [this.projectId];
      } else if (this.config.workspaceId) {
        // If workspace is provided but no project, we'll need to find or create a default project
        // For now, we'll just use the workspace
        taskData.workspace = this.config.workspaceId;
      }

      // Add assignee if email is provided
      if (assigneeEmail && (this.config.workspaceId || this.projectId)) {
        try {
          const workspaceId = this.config.workspaceId || this.projectId!;
          const users = await this.client.users.findAll({
            workspace: workspaceId,
            opt_fields: 'email',
          });
          const assignee = users.data.find((u: any) => u.email === assigneeEmail);
          if (assignee) {
            taskData.assignee = assignee.gid;
          }
        } catch (error) {
          console.warn('Could not find assignee by email, creating task without assignee:', error);
        }
      }

      const task = await this.client.tasks.create(taskData, {
        opt_fields: 'gid,name,notes',
      });

      console.log(`Created Asana task: ${task.gid}`);
      return task.gid;
    } catch (error) {
      console.error('Failed to create Asana task:', error);
      return null;
    }
  }

  async updateTask(
    taskId: string,
    context: SyncContext,
    assigneeEmail?: string
  ): Promise<boolean> {
    if (!this.client || !this.config?.enabled) {
      return false;
    }

    try {
      const updateData: any = {
        name: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
        notes: `
Content Item ID: ${context.contentItemId}
Language ID: ${context.languageId}
${context.workflowStep ? `Workflow Step: ${context.workflowStep}\n` : ''}
${context.contributors ? `Contributors: ${context.contributors.join(', ')}\n` : ''}
        `.trim(),
        due_on: context.dueDate?.toISOString().split('T')[0],
      };

      // Update assignee if email is provided
      if (assigneeEmail && this.config.workspaceId) {
        try {
          const users = await this.client!.users.findAll({
            workspace: this.config.workspaceId,
            opt_fields: 'email',
          });
          const assignee = users.data.find((u: any) => u.email === assigneeEmail);
          if (assignee) {
            updateData.assignee = assignee.gid;
          }
        } catch (error) {
          console.warn('Could not find assignee by email:', error);
        }
      }

      await this.client!.tasks.updateTask(taskId, updateData);
      console.log(`Updated Asana task: ${taskId}`);
      return true;
    } catch (error) {
      console.error('Failed to update Asana task:', error);
      return false;
    }
  }

  isEnabled(): boolean {
    return this.config?.enabled === true && this.client !== null;
  }
}

