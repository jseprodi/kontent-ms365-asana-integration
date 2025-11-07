import type { AppConfig, SyncContext } from '../types/config.js';

export class Microsoft365Service {
  private config: AppConfig['microsoft365'];

  constructor(config: AppConfig['microsoft365']) {
    this.config = config;
  }

  private getProxyBaseUrl(): string | null {
    const proxyUrl = this.config?.proxyUrl;
    if (!proxyUrl) {
      return null;
    }
    return proxyUrl.endsWith('/') ? proxyUrl.slice(0, -1) : proxyUrl;
  }

  private buildEventPayload(context: SyncContext, startTime: Date, endTime: Date) {
    return {
      subject: context.title || `Kontent.ai Content Item: ${context.contentItemId}`,
      body: {
        contentType: 'HTML',
        content: `
          <p>Content Item ID: ${context.contentItemId}</p>
          <p>Language ID: ${context.languageId}</p>
          ${context.workflowStep ? `<p>Workflow Step: ${context.workflowStep}</p>` : ''}
          ${context.contributors ? `<p>Contributors: ${context.contributors.join(', ')}</p>` : ''}
        `,
      },
      start: {
        dateTime: startTime.toISOString(),
        timeZone: 'UTC',
      },
      end: {
        dateTime: endTime.toISOString(),
        timeZone: 'UTC',
      },
      isReminderOn: true,
      reminderMinutesBeforeStart: 15,
    };
  }

  private async callProxy(
    userPrincipalName: string,
    eventPayload: any,
    eventId?: string
  ): Promise<any> {
    const proxyBase = this.getProxyBaseUrl();
    if (!proxyBase) {
      throw new Error('Microsoft 365 proxy URL is not configured.');
    }

    const proxyEndpoint = `${proxyBase}/.netlify/functions/ms365`;

    console.log('[Microsoft365Service] Calling proxy', {
      proxyEndpoint,
      hasEventId: !!eventId,
      userPrincipalName,
    });

    const response = await fetch(proxyEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        userPrincipalName,
        eventPayload,
        eventId,
      }),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error('[Microsoft365Service] Proxy call failed', {
        status: response.status,
        statusText: response.statusText,
        errorBody,
      });
      throw new Error(`Microsoft 365 proxy error: ${response.status} ${errorBody}`);
    }

    const data = await response.json();
    console.log('[Microsoft365Service] Proxy call succeeded', {
      hasId: !!data?.id,
      hasWebLink: !!data?.webLink,
    });
    return data;
  }

  async createCalendarEvent(
    userPrincipalName: string,
    context: SyncContext,
    startTime: Date,
    endTime: Date
  ): Promise<string | null> {
    console.log('[Microsoft365Service] createCalendarEvent called', {
      userPrincipalName,
      enabled: this.config?.enabled,
      hasProxyUrl: !!this.config?.proxyUrl,
      contentItemId: context.contentItemId,
    });

    if (!this.isEnabled()) {
      console.warn('[Microsoft365Service] Service is not enabled or proxy URL missing', {
        enabled: this.config?.enabled,
        proxyUrl: this.config?.proxyUrl,
      });
      return null;
    }

    try {
      const eventPayload = this.buildEventPayload(context, startTime, endTime);
      const result = await this.callProxy(userPrincipalName, eventPayload);
      if (result?.id) {
        console.log('[Microsoft365Service] Calendar event created via proxy', {
          eventId: result.id,
          subject: result.subject,
          webLink: result.webLink,
        });
        return result.id;
      }
      console.warn('[Microsoft365Service] Proxy response did not include event id');
      return null;
    } catch (error) {
      console.error('[Microsoft365Service] Failed to create calendar event via proxy:', error);
      return null;
    }
  }

  async updateCalendarEvent(
    userPrincipalName: string,
    eventId: string,
    context: SyncContext,
    startTime: Date,
    endTime: Date
  ): Promise<boolean> {
    if (!this.isEnabled()) {
      console.warn('[Microsoft365Service] Update skipped - service not enabled or proxy missing');
      return false;
    }

    try {
      const eventPayload = this.buildEventPayload(context, startTime, endTime);
      await this.callProxy(userPrincipalName, eventPayload, eventId);
      console.log('[Microsoft365Service] Calendar event updated via proxy', { eventId });
      return true;
    } catch (error) {
      console.error('[Microsoft365Service] Failed to update calendar event via proxy:', error);
      return false;
    }
  }

  isEnabled(): boolean {
    return this.config?.enabled === true && !!this.getProxyBaseUrl();
  }
}

