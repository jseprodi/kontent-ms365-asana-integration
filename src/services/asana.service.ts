import type { AppConfig, SyncContext } from '../types/config.js';

export class AsanaService {
  private config: AppConfig['asana'];
  private projectId: string | null = null;
  private readonly apiBaseUrl = 'https://app.asana.com/api/1.0';

  constructor(config: AppConfig['asana']) {
    this.config = config;
    if (config?.projectId) {
      this.projectId = config.projectId;
    }
  }

  private async makeRequest(endpoint: string, options: RequestInit = {}): Promise<any> {
    if (!this.config?.accessToken) {
      throw new Error('Asana access token is not configured');
    }

    const response = await fetch(`${this.apiBaseUrl}${endpoint}`, {
      ...options,
      headers: {
        'Authorization': `Bearer ${this.config.accessToken}`,
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: response.statusText }));
      throw new Error(`Asana API error: ${error.message || response.statusText}`);
    }

    const data = await response.json();
    return data.data || data;
  }

  private async findUserByEmail(email: string, workspaceId?: string): Promise<string | null> {
    if (!workspaceId && !this.projectId) {
      return null;
    }

    try {
      // First, get workspace ID if we only have project ID
      let actualWorkspaceId = workspaceId;
      if (!actualWorkspaceId && this.projectId) {
        const project = await this.makeRequest(`/projects/${this.projectId}`, {
          method: 'GET',
        });
        actualWorkspaceId = project.workspace?.gid;
      }

      if (!actualWorkspaceId) {
        return null;
      }

      // Get users in workspace
      const users = await this.makeRequest(`/workspaces/${actualWorkspaceId}/users`, {
        method: 'GET',
      });

      // Find user by email
      const user = Array.isArray(users)
        ? users.find((u: any) => u.email === email)
        : null;

      return user?.gid || null;
    } catch (error) {
      console.warn('Could not find user by email:', error);
      return null;
    }
  }

  async createTask(context: SyncContext, assigneeEmail?: string): Promise<string | null> {
    if (!this.config?.enabled || !this.config.accessToken) {
      console.warn('Asana service is not enabled or not initialized');
      return null;
    }

    try {
      const taskData: any = {
        data: {
          name: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
          notes: `
Content Item ID: ${context.contentItemId}
Language ID: ${context.languageId}
${context.workflowStep ? `Workflow Step: ${context.workflowStep}\n` : ''}
${context.contributors ? `Contributors: ${context.contributors.join(', ')}\n` : ''}
          `.trim(),
        },
      };

      // Add due date if available
      if (context.dueDate) {
        taskData.data.due_on = context.dueDate.toISOString().split('T')[0]; // YYYY-MM-DD format
      }

      // Add project if configured
      if (this.projectId) {
        taskData.data.projects = [this.projectId];
      } else if (this.config.workspaceId) {
        taskData.data.workspace = this.config.workspaceId;
      }

      // Add assignee if email is provided
      if (assigneeEmail) {
        const assigneeGid = await this.findUserByEmail(assigneeEmail, this.config.workspaceId);
        if (assigneeGid) {
          taskData.data.assignee = assigneeGid;
        }
      }

      const task = await this.makeRequest('/tasks', {
        method: 'POST',
        body: JSON.stringify(taskData),
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
    if (!this.config?.enabled || !this.config.accessToken) {
      return false;
    }

    try {
      const updateData: any = {
        data: {
          name: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
          notes: `
Content Item ID: ${context.contentItemId}
Language ID: ${context.languageId}
${context.workflowStep ? `Workflow Step: ${context.workflowStep}\n` : ''}
${context.contributors ? `Contributors: ${context.contributors.join(', ')}\n` : ''}
          `.trim(),
        },
      };

      // Add due date if available
      if (context.dueDate) {
        updateData.data.due_on = context.dueDate.toISOString().split('T')[0];
      }

      // Update assignee if email is provided
      if (assigneeEmail) {
        const assigneeGid = await this.findUserByEmail(assigneeEmail, this.config.workspaceId);
        if (assigneeGid) {
          updateData.data.assignee = assigneeGid;
        }
      }

      await this.makeRequest(`/tasks/${taskId}`, {
        method: 'PUT',
        body: JSON.stringify(updateData),
      });

      console.log(`Updated Asana task: ${taskId}`);
      return true;
    } catch (error) {
      console.error('Failed to update Asana task:', error);
      return false;
    }
  }

  isEnabled(): boolean {
    return this.config?.enabled === true && !!this.config.accessToken;
  }
}
