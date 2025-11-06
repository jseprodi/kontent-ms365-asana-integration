export interface AppConfig {
  microsoft365?: {
    clientId: string;
    tenantId: string;
    clientSecret: string;
    enabled: boolean;
  };
  asana?: {
    accessToken: string;
    workspaceId?: string;
    projectId?: string;
    enabled: boolean;
  };
  syncSettings?: {
    syncContributors: boolean;
    syncWorkflowSteps: boolean;
    createCalendarEvents: boolean;
    createTasks: boolean;
  };
}

export interface SyncContext {
  contentItemId: string;
  languageId: string;
  workflowStep?: string;
  contributors?: string[];
  dueDate?: Date;
  title?: string;
}

