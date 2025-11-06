import { Client } from '@microsoft/microsoft-graph-client';
import type { AppConfig, SyncContext } from '../types/config.js';

export class Microsoft365Service {
  private client: Client | null = null;
  private config: AppConfig['microsoft365'];

  constructor(config: AppConfig['microsoft365']) {
    this.config = config;
    if (config?.enabled && config.clientId && config.tenantId && config.clientSecret) {
      this.initializeClient();
    }
  }

  private async initializeClient() {
    if (!this.config) return;

    try {
      // Get access token using client credentials flow
      const tokenEndpoint = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`;
      const tokenResponse = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams({
          client_id: this.config.clientId,
          client_secret: this.config.clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      });

      if (!tokenResponse.ok) {
        throw new Error(`Failed to get access token: ${tokenResponse.statusText}`);
      }

      const tokenData = await tokenResponse.json();
      const accessToken = tokenData.access_token;

      // Initialize Graph client
      this.client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });
    } catch (error) {
      console.error('Failed to initialize Microsoft 365 client:', error);
      throw error;
    }
  }

  async createCalendarEvent(
    userId: string,
    context: SyncContext,
    startTime: Date,
    endTime: Date
  ): Promise<string | null> {
    if (!this.client || !this.config?.enabled) {
      console.warn('Microsoft 365 service is not enabled or not initialized');
      return null;
    }

    try {
      const event = {
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

      const createdEvent = await this.client!
        .api(`/users/${userId}/calendar/events`)
        .post(event);

      console.log(`Created calendar event: ${createdEvent.id}`);
      return createdEvent.id;
    } catch (error) {
      console.error('Failed to create calendar event:', error);
      return null;
    }
  }

  async updateCalendarEvent(
    userId: string,
    eventId: string,
    context: SyncContext,
    startTime: Date,
    endTime: Date
  ): Promise<boolean> {
    if (!this.client || !this.config?.enabled) {
      return false;
    }

    try {
      const event = {
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
      };

      await this.client!.api(`/users/${userId}/calendar/events/${eventId}`).patch(event);
      console.log(`Updated calendar event: ${eventId}`);
      return true;
    } catch (error) {
      console.error('Failed to update calendar event:', error);
      return false;
    }
  }

  isEnabled(): boolean {
    return this.config?.enabled === true && this.client !== null;
  }
}

