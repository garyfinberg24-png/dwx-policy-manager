// @ts-nocheck
/**
 * CalendarActionHandler
 * Handles Microsoft Graph Calendar operations within workflow execution
 * - Create calendar events (e.g., exit interviews, onboarding meetings)
 * - Update and delete calendar events
 * - Create Teams meetings
 *
 * Requires Graph API permissions:
 * - Calendars.ReadWrite (for calendar operations)
 * - OnlineMeetings.ReadWrite (for Teams meeting creation)
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  IActionContext,
  IActionResult,
  IActionConfig
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

/**
 * Calendar event details
 */
export interface ICalendarEventDetails {
  id?: string;
  subject: string;
  body?: string;
  start: Date;
  end: Date;
  location?: string;
  attendees?: string[];
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
}

/**
 * Result of calendar operation
 */
export interface ICalendarOperationResult {
  success: boolean;
  eventId?: string;
  onlineMeetingUrl?: string;
  webLink?: string;
  error?: string;
}

export class CalendarActionHandler {
  private context: WebPartContext;
  private siteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // ============================================================================
  // EVENT CREATION
  // ============================================================================

  /**
   * Create a calendar event
   */
  public async createCalendarEvent(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const calendarUserId = this.resolveCalendarUserId(config, context);
      const eventDetails = this.buildEventDetails(config, context);

      if (!eventDetails.subject) {
        return { success: false, error: 'Event title/subject is required' };
      }

      const graphClient = await this.getGraphClient();

      // Build the event payload
      const eventPayload: Record<string, unknown> = {
        subject: eventDetails.subject,
        body: eventDetails.body ? {
          contentType: 'HTML',
          content: eventDetails.body
        } : undefined,
        start: {
          dateTime: eventDetails.start.toISOString(),
          timeZone: 'UTC'
        },
        end: {
          dateTime: eventDetails.end.toISOString(),
          timeZone: 'UTC'
        },
        location: eventDetails.location ? {
          displayName: eventDetails.location
        } : undefined,
        attendees: this.buildAttendeesPayload(eventDetails.attendees),
        isOnlineMeeting: eventDetails.isOnlineMeeting || false,
        onlineMeetingProvider: eventDetails.isOnlineMeeting ? 'teamsForBusiness' : undefined
      };

      // Create event - use /me/events or /users/{id}/events based on context
      let apiPath: string;
      if (calendarUserId && calendarUserId !== 'me') {
        apiPath = `/users/${calendarUserId}/events`;
      } else {
        apiPath = '/me/events';
      }

      const createdEvent = await graphClient
        .api(apiPath)
        .post(eventPayload);

      logger.info('CalendarActionHandler', `Created calendar event: ${createdEvent.id}`, {
        subject: eventDetails.subject,
        start: eventDetails.start.toISOString()
      });

      return {
        success: true,
        outputVariables: {
          eventId: createdEvent.id,
          eventWebLink: createdEvent.webLink,
          onlineMeetingUrl: createdEvent.onlineMeeting?.joinUrl,
          eventStart: eventDetails.start.toISOString(),
          eventEnd: eventDetails.end.toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('CalendarActionHandler', 'Error creating calendar event', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create calendar event'
      };
    }
  }

  /**
   * Create an exit interview event
   * Specialized method for leaver workflow
   */
  public async createExitInterview(
    employeeEmail: string,
    managerEmail: string,
    hrEmail: string,
    lastWorkingDay: Date,
    employeeName: string,
    processId: number
  ): Promise<ICalendarOperationResult> {
    try {
      const graphClient = await this.getGraphClient();

      // Schedule exit interview 2 days before last working day, or today if last day is soon
      const now = new Date();
      const twoDaysBefore = new Date(lastWorkingDay);
      twoDaysBefore.setDate(twoDaysBefore.getDate() - 2);

      const interviewDate = twoDaysBefore > now ? twoDaysBefore : now;
      interviewDate.setHours(14, 0, 0, 0); // 2 PM

      const endDate = new Date(interviewDate);
      endDate.setMinutes(endDate.getMinutes() + 60); // 1 hour meeting

      const eventPayload = {
        subject: `Exit Interview: ${employeeName}`,
        body: {
          contentType: 'HTML',
          content: `
            <p>Exit interview for <strong>${employeeName}</strong></p>
            <p>Last working day: ${lastWorkingDay.toLocaleDateString()}</p>
            <p>Process ID: #${processId}</p>
            <hr>
            <p>Please use this time to discuss:</p>
            <ul>
              <li>Reason for leaving</li>
              <li>Feedback on role and department</li>
              <li>Knowledge transfer status</li>
              <li>Outstanding items or concerns</li>
            </ul>
            <p><a href="${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${processId}">View Process Details</a></p>
          `
        },
        start: {
          dateTime: interviewDate.toISOString(),
          timeZone: 'UTC'
        },
        end: {
          dateTime: endDate.toISOString(),
          timeZone: 'UTC'
        },
        attendees: [
          { emailAddress: { address: employeeEmail }, type: 'required' },
          { emailAddress: { address: managerEmail }, type: 'required' },
          { emailAddress: { address: hrEmail }, type: 'required' }
        ],
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness'
      };

      const createdEvent = await graphClient
        .api('/me/events')
        .post(eventPayload);

      logger.info('CalendarActionHandler', `Created exit interview event: ${createdEvent.id}`);

      return {
        success: true,
        eventId: createdEvent.id,
        onlineMeetingUrl: createdEvent.onlineMeeting?.joinUrl,
        webLink: createdEvent.webLink
      };
    } catch (error) {
      logger.error('CalendarActionHandler', 'Error creating exit interview', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create exit interview'
      };
    }
  }

  /**
   * Create an onboarding welcome meeting
   * Specialized method for joiner workflow
   */
  public async createWelcomeMeeting(
    employeeEmail: string,
    managerEmail: string,
    buddyEmail: string | undefined,
    startDate: Date,
    employeeName: string,
    processId: number
  ): Promise<ICalendarOperationResult> {
    try {
      const graphClient = await this.getGraphClient();

      // Schedule welcome meeting on start date at 10 AM
      const meetingDate = new Date(startDate);
      meetingDate.setHours(10, 0, 0, 0);

      const endDate = new Date(meetingDate);
      endDate.setMinutes(endDate.getMinutes() + 30); // 30 minute meeting

      const attendees = [
        { emailAddress: { address: employeeEmail }, type: 'required' },
        { emailAddress: { address: managerEmail }, type: 'required' }
      ];

      if (buddyEmail) {
        attendees.push({ emailAddress: { address: buddyEmail }, type: 'optional' });
      }

      const eventPayload = {
        subject: `Welcome Meeting: ${employeeName}`,
        body: {
          contentType: 'HTML',
          content: `
            <p>Welcome to the team, <strong>${employeeName}</strong>!</p>
            <p>This is your first day welcome meeting with your manager${buddyEmail ? ' and buddy' : ''}.</p>
            <hr>
            <p>Agenda:</p>
            <ul>
              <li>Team introductions</li>
              <li>First week overview</li>
              <li>IT setup and access verification</li>
              <li>Questions and answers</li>
            </ul>
            <p><a href="${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${processId}">View Onboarding Progress</a></p>
          `
        },
        start: {
          dateTime: meetingDate.toISOString(),
          timeZone: 'UTC'
        },
        end: {
          dateTime: endDate.toISOString(),
          timeZone: 'UTC'
        },
        attendees,
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness'
      };

      const createdEvent = await graphClient
        .api('/me/events')
        .post(eventPayload);

      logger.info('CalendarActionHandler', `Created welcome meeting event: ${createdEvent.id}`);

      return {
        success: true,
        eventId: createdEvent.id,
        onlineMeetingUrl: createdEvent.onlineMeeting?.joinUrl,
        webLink: createdEvent.webLink
      };
    } catch (error) {
      logger.error('CalendarActionHandler', 'Error creating welcome meeting', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create welcome meeting'
      };
    }
  }

  // ============================================================================
  // EVENT UPDATES
  // ============================================================================

  /**
   * Update an existing calendar event
   */
  public async updateCalendarEvent(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      // Get event ID from context variables (set by previous createCalendarEvent)
      const eventId = context.variables['eventId'] as string || config.itemId?.toString();

      if (!eventId) {
        return { success: false, error: 'Event ID not specified' };
      }

      const calendarUserId = this.resolveCalendarUserId(config, context);
      const graphClient = await this.getGraphClient();

      // Build update payload from config
      const updatePayload: Record<string, unknown> = {};

      if (config.eventTitle) {
        updatePayload.subject = this.resolveTemplateString(config.eventTitle, context);
      }

      if (config.eventDescription) {
        updatePayload.body = {
          contentType: 'HTML',
          content: this.resolveTemplateString(config.eventDescription, context)
        };
      }

      if (config.eventLocation) {
        updatePayload.location = { displayName: config.eventLocation };
      }

      if (Object.keys(updatePayload).length === 0) {
        return { success: false, error: 'No updates specified' };
      }

      // Update event
      let apiPath: string;
      if (calendarUserId && calendarUserId !== 'me') {
        apiPath = `/users/${calendarUserId}/events/${eventId}`;
      } else {
        apiPath = `/me/events/${eventId}`;
      }

      await graphClient
        .api(apiPath)
        .patch(updatePayload);

      logger.info('CalendarActionHandler', `Updated calendar event: ${eventId}`);

      return {
        success: true,
        outputVariables: {
          eventId,
          updatedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('CalendarActionHandler', 'Error updating calendar event', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update calendar event'
      };
    }
  }

  /**
   * Delete/cancel a calendar event
   */
  public async deleteCalendarEvent(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const eventId = context.variables['eventId'] as string || config.itemId?.toString();

      if (!eventId) {
        return { success: false, error: 'Event ID not specified' };
      }

      const calendarUserId = this.resolveCalendarUserId(config, context);
      const graphClient = await this.getGraphClient();

      let apiPath: string;
      if (calendarUserId && calendarUserId !== 'me') {
        apiPath = `/users/${calendarUserId}/events/${eventId}`;
      } else {
        apiPath = `/me/events/${eventId}`;
      }

      await graphClient
        .api(apiPath)
        .delete();

      logger.info('CalendarActionHandler', `Deleted calendar event: ${eventId}`);

      return {
        success: true,
        outputVariables: {
          deletedEventId: eventId,
          deletedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('CalendarActionHandler', 'Error deleting calendar event', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to delete calendar event'
      };
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private async getGraphClient(): Promise<MSGraphClientV3> {
    return this.context.msGraphClientFactory.getClient('3');
  }

  private resolveCalendarUserId(config: IActionConfig, context: IActionContext): string | undefined {
    if (config.calendarUserId) {
      return config.calendarUserId;
    }

    if (config.calendarUserIdField) {
      const fieldValue = context.process[config.calendarUserIdField];
      if (typeof fieldValue === 'string') {
        return fieldValue;
      }
    }

    return undefined; // Will use /me endpoint
  }

  private buildEventDetails(config: IActionConfig, context: IActionContext): ICalendarEventDetails {
    // Resolve title
    let subject = config.eventTitle || '';
    if (config.eventTitleTemplate) {
      subject = this.resolveTemplateString(config.eventTitleTemplate, context);
    }

    // Resolve description
    let body = config.eventDescription;
    if (config.eventDescriptionTemplate) {
      body = this.resolveTemplateString(config.eventDescriptionTemplate, context);
    }

    // Resolve dates
    const start = this.resolveDate(config, context, 'start');
    const end = this.resolveDate(config, context, 'end') || this.addMinutes(start, config.eventDuration || 60);

    // Resolve attendees
    let attendees: string[] = [];
    if (config.eventAttendees) {
      attendees = [...config.eventAttendees];
    }
    if (config.eventAttendeesField) {
      const fieldValue = context.process[config.eventAttendeesField];
      if (typeof fieldValue === 'string') {
        attendees.push(fieldValue);
      } else if (Array.isArray(fieldValue)) {
        attendees.push(...fieldValue.filter((v): v is string => typeof v === 'string'));
      }
    }

    return {
      subject,
      body,
      start,
      end,
      location: config.eventLocation,
      attendees,
      isOnlineMeeting: config.eventIsOnline
    };
  }

  private resolveDate(config: IActionConfig, context: IActionContext, type: 'start' | 'end'): Date {
    const dateField = type === 'start' ? config.eventStartDateField : config.eventEndDateField;
    const dateValue = type === 'start' ? config.eventStartDate : config.eventEndDate;

    // From field
    if (dateField) {
      const fieldValue = context.process[dateField];
      if (fieldValue instanceof Date) {
        return fieldValue;
      }
      if (typeof fieldValue === 'string') {
        return new Date(fieldValue);
      }
    }

    // Static value
    if (dateValue) {
      return new Date(dateValue);
    }

    // Default to tomorrow for start, start + 1 hour for end
    if (type === 'start') {
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      tomorrow.setHours(9, 0, 0, 0);
      return tomorrow;
    }

    return new Date();
  }

  private addMinutes(date: Date, minutes: number): Date {
    return new Date(date.getTime() + minutes * 60000);
  }

  private resolveTemplateString(template: string, context: IActionContext): string {
    let result = template;

    // Replace {{fieldName}} placeholders with process values
    const placeholderRegex = /\{\{(\w+)\}\}/g;
    result = result.replace(placeholderRegex, (match, fieldName) => {
      const value = context.process[fieldName] || context.variables[fieldName];
      if (value !== undefined && value !== null) {
        return String(value);
      }
      return match; // Keep original if not found
    });

    return result;
  }

  private buildAttendeesPayload(attendees: string[] | undefined): Array<{ emailAddress: { address: string }; type: string }> | undefined {
    if (!attendees || attendees.length === 0) {
      return undefined;
    }

    return attendees.map(email => ({
      emailAddress: { address: email },
      type: 'required'
    }));
  }
}

export default CalendarActionHandler;
