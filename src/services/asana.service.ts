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

    const url = `${this.apiBaseUrl}${endpoint}`;
    console.log('[AsanaService] Making API request', {
      method: options.method || 'GET',
      url,
      hasBody: !!options.body,
    });

    const response = await fetch(url, {
      ...options,
      headers: {
        'Authorization': `Bearer ${this.config.accessToken}`,
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    console.log('[AsanaService] API response received', {
      status: response.status,
      statusText: response.statusText,
      ok: response.ok,
    });

    if (!response.ok) {
      const errorText = await response.text();
      let error;
      try {
        error = JSON.parse(errorText);
      } catch {
        error = { message: errorText || response.statusText };
      }
      console.error('[AsanaService] API error response', {
        status: response.status,
        error,
      });
      throw new Error(`Asana API error: ${error.message || error.errors?.[0]?.message || response.statusText}`);
    }

    const data = await response.json();
    console.log('[AsanaService] API response parsed', {
      hasData: !!data.data,
      dataKeys: data.data ? Object.keys(data.data) : Object.keys(data),
    });
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
    console.log('[AsanaService] createTask called', {
      enabled: this.config?.enabled,
      hasAccessToken: !!this.config?.accessToken,
      contentItemId: context.contentItemId,
      assigneeEmail,
    });

    if (!this.config?.enabled || !this.config.accessToken) {
      console.warn('[AsanaService] Service is not enabled or not initialized', {
        enabled: this.config?.enabled,
        hasAccessToken: !!this.config?.accessToken,
      });
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
        console.log('[AsanaService] Added due date', { due_on: taskData.data.due_on });
      }

      // Add project if configured
      if (this.projectId) {
        taskData.data.projects = [this.projectId];
        console.log('[AsanaService] Added project', { projectId: this.projectId });
      } else if (this.config.workspaceId) {
        taskData.data.workspace = this.config.workspaceId;
        console.log('[AsanaService] Added workspace', { workspaceId: this.config.workspaceId });
      }

      // Add assignee if email is provided
      if (assigneeEmail) {
        console.log('[AsanaService] Finding user by email', { email: assigneeEmail });
        const assigneeGid = await this.findUserByEmail(assigneeEmail, this.config.workspaceId);
        if (assigneeGid) {
          taskData.data.assignee = assigneeGid;
          console.log('[AsanaService] Added assignee', { assigneeGid });
        } else {
          console.warn('[AsanaService] Could not find assignee by email', { email: assigneeEmail });
        }
      }

      console.log('[AsanaService] Creating task', {
        endpoint: '/tasks',
        taskData: {
          name: taskData.data.name,
          hasProject: !!taskData.data.projects,
          hasWorkspace: !!taskData.data.workspace,
          hasAssignee: !!taskData.data.assignee,
          hasDueDate: !!taskData.data.due_on,
        },
      });

      const task = await this.makeRequest('/tasks', {
        method: 'POST',
        body: JSON.stringify(taskData),
      });

      console.log('[AsanaService] Task created successfully', {
        taskId: task.gid,
        name: task.name,
        html_url: task.permalink_url,
      });
      return task.gid;
    } catch (error: any) {
      console.error('[AsanaService] Failed to create task:', {
        error: error.message,
        stack: error.stack,
        response: error.response,
      });
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
