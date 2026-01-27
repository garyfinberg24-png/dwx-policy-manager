// @ts-nocheck
// Calendar Service
// Microsoft Graph Calendar integration for JML events
// Handles scheduling, meeting rooms, and availability

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/calendars';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IGraphCalendarEvent,
  IGraphDateTimeZone,
  IGraphAttendee,
  IGraphMeetingRoom,
  IRoomAvailability,
  IScheduleRequest,
  IScheduleResponse,
  IMeetingTimeSuggestion,
  IScheduleEventRequest,
  IScheduleEventResult,
  IExitInterviewRequest,
  IExitInterviewResult,
  IOnboardingSchedule,
  IOnboardingEvent,
  IOnboardingTemplate,
  GraphEventType,
  EventStatus,
  MeetingType
} from '../models/ICalendar';
import { IJmlProcess } from '../models';
import { logger } from './LoggingService';

export class CalendarService {
  private sp: SPFI;
  private graph: GraphFI;
  private context: WebPartContext;
  private defaultTimeZone: string = 'UTC';

  private readonly CALENDAR_EVENTS_LIST = 'JML_CalendarEvents';
  private readonly EVENT_TEMPLATES_LIST = 'JML_EventTemplates';
  private readonly ONBOARDING_TEMPLATES_LIST = 'JML_OnboardingTemplates';

  constructor(sp: SPFI, graph: GraphFI, context: WebPartContext) {
    this.sp = sp;
    this.graph = graph;
    this.context = context;
  }

  // ============================================================================
  // Calendar Event Operations
  // ============================================================================

  /**
   * Create a calendar event via Microsoft Graph
   */
  public async createEvent(
    event: IGraphCalendarEvent,
    userEmail?: string
  ): Promise<IGraphCalendarEvent> {
    try {
      const graphEvent = this.mapToGraphEvent(event);

      let createdEvent: any;
      if (userEmail) {
        // Create event on specific user's calendar
        createdEvent = await (this.graph.users.getById(userEmail).calendar as any).events.add(graphEvent);
      } else {
        // Create event on current user's calendar
        createdEvent = await (this.graph.me.calendar as any).events.add(graphEvent);
      }

      logger.info('CalendarService', `Created event: ${createdEvent.subject}`);

      return this.mapFromGraphEvent(createdEvent);
    } catch (error) {
      logger.error('CalendarService', 'Error creating calendar event:', error);
      throw error;
    }
  }

  /**
   * Update an existing calendar event
   */
  public async updateEvent(
    eventId: string,
    updates: Partial<IGraphCalendarEvent>,
    userEmail?: string
  ): Promise<void> {
    try {
      const graphUpdates = this.mapToGraphEvent(updates as IGraphCalendarEvent);

      if (userEmail) {
        await (this.graph.users.getById(userEmail).calendar as any).events.getById(eventId).update(graphUpdates);
      } else {
        await (this.graph.me.calendar as any).events.getById(eventId).update(graphUpdates);
      }

      logger.info('CalendarService', `Updated event: ${eventId}`);
    } catch (error) {
      logger.error('CalendarService', 'Error updating calendar event:', error);
      throw error;
    }
  }

  /**
   * Delete a calendar event
   */
  public async deleteEvent(eventId: string, userEmail?: string): Promise<void> {
    try {
      if (userEmail) {
        await (this.graph.users.getById(userEmail).calendar as any).events.getById(eventId).delete();
      } else {
        await (this.graph.me.calendar as any).events.getById(eventId).delete();
      }

      logger.info('CalendarService', `Deleted event: ${eventId}`);
    } catch (error) {
      logger.error('CalendarService', 'Error deleting calendar event:', error);
      throw error;
    }
  }

  /**
   * Get events for a date range
   */
  public async getEvents(
    startDate: Date,
    endDate: Date,
    userEmail?: string
  ): Promise<IGraphCalendarEvent[]> {
    try {
      const startDateTime = startDate.toISOString();
      const endDateTime = endDate.toISOString();

      let events: any[];
      if (userEmail) {
        events = await (this.graph.users.getById(userEmail).calendar as any).events
          .filter(`start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`)
          .orderBy('start/dateTime')
          .top(100)();
      } else {
        events = await (this.graph.me.calendar as any).events
          .filter(`start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`)
          .orderBy('start/dateTime')
          .top(100)();
      }

      return events.map(e => this.mapFromGraphEvent(e));
    } catch (error) {
      logger.error('CalendarService', 'Error getting calendar events:', error);
      throw error;
    }
  }

  // ============================================================================
  // Meeting Room Operations
  // ============================================================================

  /**
   * Get all available meeting rooms
   */
  public async getMeetingRooms(): Promise<IGraphMeetingRoom[]> {
    try {
      const rooms = await (this.graph as any).places.rooms();
      return rooms.map((r: any) => this.mapMeetingRoom(r));
    } catch (error) {
      logger.error('CalendarService', 'Error getting meeting rooms:', error);
      return [];
    }
  }

  /**
   * Get room lists (buildings)
   */
  public async getRoomLists(): Promise<Array<{ id: string; displayName: string; emailAddress: string }>> {
    try {
      const lists = await (this.graph as any).places.roomLists();
      return lists.map((l: any) => ({
        id: l.id,
        displayName: l.displayName,
        emailAddress: l.emailAddress
      }));
    } catch (error) {
      logger.error('CalendarService', 'Error getting room lists:', error);
      return [];
    }
  }

  /**
   * Find available rooms for a time slot
   */
  public async findAvailableRooms(
    startTime: Date,
    endTime: Date,
    capacity?: number,
    building?: string
  ): Promise<IRoomAvailability[]> {
    try {
      // Get all rooms
      let rooms = await this.getMeetingRooms();

      // Filter by capacity if specified
      if (capacity) {
        rooms = rooms.filter(r => (r.capacity || 0) >= capacity);
      }

      // Filter by building if specified
      if (building) {
        rooms = rooms.filter(r => r.building?.toLowerCase() === building.toLowerCase());
      }

      // Check availability for each room
      const availability: IRoomAvailability[] = [];

      for (const room of rooms) {
        try {
          const schedules = await this.getSchedule(
            [room.emailAddress],
            startTime,
            endTime
          );

          const schedule = schedules[0];
          const isBusy = schedule?.scheduleItems.some(
            item => item.status === 'Busy' || item.status === 'Tentative'
          );

          availability.push({
            room,
            availability: isBusy ? 'Busy' : 'Available',
            availabilityView: schedule?.availabilityView
          });
        } catch {
          availability.push({
            room,
            availability: 'Unknown'
          });
        }
      }

      return availability.sort((a, b) => {
        // Sort available rooms first
        if (a.availability === 'Available' && b.availability !== 'Available') return -1;
        if (a.availability !== 'Available' && b.availability === 'Available') return 1;
        return 0;
      });
    } catch (error) {
      logger.error('CalendarService', 'Error finding available rooms:', error);
      return [];
    }
  }

  // ============================================================================
  // Schedule / Availability Operations
  // ============================================================================

  /**
   * Get schedule/availability for users or rooms
   */
  public async getSchedule(
    emails: string[],
    startTime: Date,
    endTime: Date,
    availabilityViewInterval: number = 30
  ): Promise<IScheduleResponse[]> {
    try {
      const request: IScheduleRequest = {
        schedules: emails,
        startTime: {
          dateTime: startTime.toISOString(),
          timeZone: this.defaultTimeZone
        },
        endTime: {
          dateTime: endTime.toISOString(),
          timeZone: this.defaultTimeZone
        },
        availabilityViewInterval
      };

      const response = await (this.graph.me.calendar as any).getSchedule(request);
      return response.value || [];
    } catch (error) {
      logger.error('CalendarService', 'Error getting schedule:', error);
      throw error;
    }
  }

  /**
   * Find meeting times that work for all attendees
   */
  public async findMeetingTimes(
    attendeeEmails: string[],
    startDate: Date,
    endDate: Date,
    durationMinutes: number,
    maxCandidates: number = 5
  ): Promise<IMeetingTimeSuggestion[]> {
    try {
      const request = {
        attendees: attendeeEmails.map(email => ({
          emailAddress: { address: email },
          type: 'Required' as const
        })),
        timeConstraint: {
          activityDomain: 'Work' as const,
          timeSlots: [{
            start: {
              dateTime: startDate.toISOString(),
              timeZone: this.defaultTimeZone
            },
            end: {
              dateTime: endDate.toISOString(),
              timeZone: this.defaultTimeZone
            }
          }]
        },
        meetingDuration: `PT${durationMinutes}M`,
        maxCandidates,
        isOrganizerOptional: false,
        returnSuggestionReasons: true
      };

      const response = await (this.graph.me as any).findMeetingTimes(request);
      return response.meetingTimeSuggestions || [];
    } catch (error) {
      logger.error('CalendarService', 'Error finding meeting times:', error);
      return [];
    }
  }

  // ============================================================================
  // JML-Specific Scheduling
  // ============================================================================

  /**
   * Schedule a JML-related event with smart time finding
   */
  public async scheduleJmlEvent(request: IScheduleEventRequest): Promise<IScheduleEventResult> {
    try {
      const allAttendees = [request.employeeEmail, request.organizerEmail, ...(request.additionalAttendees || [])];

      // Determine date range for scheduling
      const startDate = request.preferredDate || new Date();
      const endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 14); // Look 2 weeks out

      // Find available meeting times
      let meetingTimes = await this.findMeetingTimes(
        allAttendees,
        startDate,
        endDate,
        request.duration
      );

      // Filter by preferred time range if specified
      if (request.preferredTimeRange && meetingTimes.length > 0) {
        meetingTimes = meetingTimes.filter(mt => {
          const hour = new Date(mt.meetingTimeSlot.start.dateTime).getHours();
          return hour >= request.preferredTimeRange!.startHour &&
                 hour <= request.preferredTimeRange!.endHour;
        });
      }

      if (meetingTimes.length === 0) {
        return {
          success: false,
          errorMessage: 'No available meeting times found for all attendees',
          alternativeTimes: []
        };
      }

      // Get first available slot
      const selectedTime = meetingTimes[0];

      // Find a meeting room if needed
      let meetingRoom: IGraphMeetingRoom | undefined;
      if (request.meetingType === 'InPerson' || request.meetingType === 'Hybrid') {
        const rooms = await this.findAvailableRooms(
          new Date(selectedTime.meetingTimeSlot.start.dateTime),
          new Date(selectedTime.meetingTimeSlot.end.dateTime),
          request.roomRequirements?.capacity,
          request.roomRequirements?.building
        );

        const availableRoom = rooms.find(r => r.availability === 'Available');
        if (availableRoom) {
          meetingRoom = availableRoom.room;
        }
      }

      // Create the event
      const event = await this.createJmlEvent({
        processId: request.processId,
        eventType: request.eventType,
        subject: request.subject || this.getDefaultSubject(request.eventType, request.employeeName),
        body: request.body || this.getDefaultBody(request.eventType, request.employeeName),
        start: selectedTime.meetingTimeSlot.start,
        end: selectedTime.meetingTimeSlot.end,
        attendees: allAttendees,
        isOnlineMeeting: request.isOnlineMeeting ?? (request.meetingType !== 'InPerson'),
        meetingRoom,
        meetingType: request.meetingType
      });

      return {
        success: true,
        eventId: event.id,
        scheduledTime: selectedTime.meetingTimeSlot.start,
        meetingRoom,
        onlineMeetingUrl: event.onlineMeeting?.joinUrl,
        alternativeTimes: meetingTimes.slice(1, 4) // Include alternatives
      };
    } catch (error) {
      logger.error('CalendarService', 'Error scheduling JML event:', error);
      return {
        success: false,
        errorMessage: error instanceof Error ? error.message : 'Failed to schedule event'
      };
    }
  }

  /**
   * Schedule an exit interview for a leaver
   */
  public async scheduleExitInterview(request: IExitInterviewRequest): Promise<IExitInterviewResult> {
    try {
      const attendees = [request.employeeEmail, request.hrContactEmail];
      if (request.includeManager && request.managerEmail) {
        attendees.push(request.managerEmail);
      }

      // Schedule 2-5 days before last working day
      const preferredDate = request.preferredDate || new Date(request.lastWorkingDay);
      if (!request.preferredDate) {
        preferredDate.setDate(preferredDate.getDate() - 3);
      }

      const duration = request.duration || 60;

      const result = await this.scheduleJmlEvent({
        processId: request.processId,
        eventType: 'ExitInterview',
        employeeEmail: request.employeeEmail,
        employeeName: request.employeeName,
        organizerEmail: request.hrContactEmail,
        additionalAttendees: request.managerEmail && request.includeManager ? [request.managerEmail] : [],
        preferredDate,
        preferredTimeRange: { startHour: 9, endHour: 17 },
        duration,
        meetingType: 'Hybrid',
        isOnlineMeeting: request.isOnlineMeeting ?? true,
        subject: `Exit Interview - ${request.employeeName}`,
        body: this.buildExitInterviewBody(request)
      });

      if (result.success) {
        // Track in SharePoint
        await this.trackCalendarEvent({
          processId: request.processId,
          eventType: 'ExitInterview',
          graphEventId: result.eventId!,
          scheduledDate: new Date(result.scheduledTime!.dateTime),
          attendees: attendees.join(';'),
          status: 'Scheduled'
        });
      }

      return {
        success: result.success,
        eventId: result.eventId,
        scheduledDateTime: result.scheduledTime ? new Date(result.scheduledTime.dateTime) : undefined,
        attendees,
        meetingUrl: result.onlineMeetingUrl,
        errorMessage: result.errorMessage
      };
    } catch (error) {
      logger.error('CalendarService', 'Error scheduling exit interview:', error);
      return {
        success: false,
        attendees: [],
        errorMessage: error instanceof Error ? error.message : 'Failed to schedule exit interview'
      };
    }
  }

  /**
   * Schedule full onboarding calendar for a new joiner
   */
  public async scheduleOnboarding(
    process: IJmlProcess,
    template?: IOnboardingTemplate
  ): Promise<IOnboardingSchedule> {
    try {
      // Get default template if not provided
      if (!template) {
        template = await this.getDefaultOnboardingTemplate(process.Department);
      }

      const schedule: IOnboardingSchedule = {
        processId: process.Id!,
        employeeId: process.EmployeeID || '',
        employeeName: process.EmployeeName,
        startDate: process.StartDate,
        events: [],
        status: 'Draft',
        createdAt: new Date(),
        modifiedAt: new Date()
      };

      // Schedule each event from the template
      for (const templateEvent of (template?.events || [])) {
        const eventDate = new Date(process.StartDate);
        eventDate.setDate(eventDate.getDate() + templateEvent.dayOffset);

        const onboardingEvent: IOnboardingEvent = {
          ...templateEvent,
          status: EventStatus.Scheduled
        };

        try {
          const result = await this.scheduleJmlEvent({
            processId: process.Id!,
            eventType: templateEvent.eventType,
            employeeEmail: process.EmployeeEmail,
            employeeName: process.EmployeeName,
            organizerEmail: templateEvent.organizer,
            additionalAttendees: templateEvent.attendees,
            preferredDate: eventDate,
            preferredTimeRange: { startHour: 9, endHour: 17 },
            duration: templateEvent.duration,
            meetingType: templateEvent.isOnlineMeeting ? 'Online' : 'InPerson',
            isOnlineMeeting: templateEvent.isOnlineMeeting,
            subject: templateEvent.subject.replace('{EmployeeName}', process.EmployeeName)
          });

          if (result.success) {
            onboardingEvent.scheduledEventId = result.eventId;
            onboardingEvent.status = EventStatus.Scheduled;
          } else {
            onboardingEvent.status = EventStatus.Pending;
          }
        } catch {
          onboardingEvent.status = EventStatus.Pending;
        }

        schedule.events.push(onboardingEvent);
      }

      schedule.status = schedule.events.every(e => e.status === EventStatus.Scheduled)
        ? 'Scheduled'
        : 'Draft';

      // Save schedule to SharePoint
      await this.saveOnboardingSchedule(schedule);

      return schedule;
    } catch (error) {
      logger.error('CalendarService', 'Error scheduling onboarding:', error);
      throw error;
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  private async createJmlEvent(params: {
    processId: number;
    eventType: GraphEventType;
    subject: string;
    body: string;
    start: IGraphDateTimeZone;
    end: IGraphDateTimeZone;
    attendees: string[];
    isOnlineMeeting: boolean;
    meetingRoom?: IGraphMeetingRoom;
    meetingType: MeetingType;
  }): Promise<IGraphCalendarEvent> {
    const attendeesList: IGraphAttendee[] = params.attendees.map(email => ({
      emailAddress: { address: email },
      type: 'Required' as const
    }));

    // Add meeting room as resource if specified
    if (params.meetingRoom) {
      attendeesList.push({
        emailAddress: {
          address: params.meetingRoom.emailAddress,
          name: params.meetingRoom.displayName
        },
        type: 'Resource'
      });
    }

    const event: IGraphCalendarEvent = {
      subject: params.subject,
      body: {
        contentType: 'HTML',
        content: params.body
      },
      start: params.start,
      end: params.end,
      attendees: attendeesList,
      isOnlineMeeting: params.isOnlineMeeting,
      onlineMeetingProvider: params.isOnlineMeeting ? 'TeamsForBusiness' : undefined,
      location: params.meetingRoom ? {
        displayName: params.meetingRoom.displayName,
        locationType: 'ConferenceRoom',
        locationEmailAddress: params.meetingRoom.emailAddress
      } : undefined,
      showAs: 'Busy',
      importance: 'Normal',
      reminderMinutesBeforeStart: 15,
      responseRequested: true,
      categories: ['JML', params.eventType]
    };

    return this.createEvent(event);
  }

  private mapToGraphEvent(event: IGraphCalendarEvent): any {
    const graphEvent: any = {};

    if (event.subject) graphEvent.subject = event.subject;
    if (event.body) graphEvent.body = event.body;
    if (event.start) graphEvent.start = event.start;
    if (event.end) graphEvent.end = event.end;
    if (event.location) graphEvent.location = event.location;
    if (event.attendees) graphEvent.attendees = event.attendees;
    if (event.isOnlineMeeting !== undefined) graphEvent.isOnlineMeeting = event.isOnlineMeeting;
    if (event.onlineMeetingProvider) graphEvent.onlineMeetingProvider = event.onlineMeetingProvider;
    if (event.showAs) graphEvent.showAs = event.showAs;
    if (event.importance) graphEvent.importance = event.importance;
    if (event.sensitivity) graphEvent.sensitivity = event.sensitivity;
    if (event.categories) graphEvent.categories = event.categories;
    if (event.isAllDay !== undefined) graphEvent.isAllDay = event.isAllDay;
    if (event.reminderMinutesBeforeStart !== undefined) {
      graphEvent.reminderMinutesBeforeStart = event.reminderMinutesBeforeStart;
    }
    if (event.responseRequested !== undefined) graphEvent.responseRequested = event.responseRequested;

    return graphEvent;
  }

  private mapFromGraphEvent(graphEvent: any): IGraphCalendarEvent {
    return {
      id: graphEvent.id,
      subject: graphEvent.subject,
      bodyPreview: graphEvent.bodyPreview,
      body: graphEvent.body,
      start: graphEvent.start,
      end: graphEvent.end,
      location: graphEvent.location,
      attendees: graphEvent.attendees || [],
      organizer: graphEvent.organizer,
      isOnlineMeeting: graphEvent.isOnlineMeeting,
      onlineMeetingProvider: graphEvent.onlineMeetingProvider,
      onlineMeeting: graphEvent.onlineMeeting,
      showAs: graphEvent.showAs,
      importance: graphEvent.importance,
      sensitivity: graphEvent.sensitivity,
      categories: graphEvent.categories,
      isAllDay: graphEvent.isAllDay,
      isCancelled: graphEvent.isCancelled,
      reminderMinutesBeforeStart: graphEvent.reminderMinutesBeforeStart,
      responseRequested: graphEvent.responseRequested,
      webLink: graphEvent.webLink,
      createdDateTime: graphEvent.createdDateTime ? new Date(graphEvent.createdDateTime) : undefined,
      lastModifiedDateTime: graphEvent.lastModifiedDateTime ? new Date(graphEvent.lastModifiedDateTime) : undefined
    };
  }

  private mapMeetingRoom(room: any): IGraphMeetingRoom {
    return {
      id: room.id,
      emailAddress: room.emailAddress,
      displayName: room.displayName,
      capacity: room.capacity,
      building: room.building,
      floorNumber: room.floorNumber,
      floorLabel: room.floorLabel,
      isWheelChairAccessible: room.isWheelChairAccessible,
      bookingType: room.bookingType,
      tags: room.tags,
      address: room.address
    };
  }

  private getDefaultSubject(eventType: GraphEventType, employeeName: string): string {
    const subjects: Record<GraphEventType, string> = {
      ExitInterview: `Exit Interview - ${employeeName}`,
      OnboardingSession: `Onboarding Session - ${employeeName}`,
      TrainingSession: `Training Session - ${employeeName}`,
      EquipmentHandover: `Equipment Handover - ${employeeName}`,
      AccessReview: `Access Review - ${employeeName}`,
      TeamIntroduction: `Team Introduction - ${employeeName}`,
      HRMeeting: `HR Meeting - ${employeeName}`,
      ITSetup: `IT Setup - ${employeeName}`,
      SecurityBriefing: `Security Briefing - ${employeeName}`,
      PolicyReview: `Policy Review - ${employeeName}`,
      ManagerMeeting: `Manager Meeting - ${employeeName}`,
      BuddyIntroduction: `Buddy Introduction - ${employeeName}`,
      DepartmentOrientation: `Department Orientation - ${employeeName}`,
      Custom: `Meeting - ${employeeName}`
    };
    return subjects[eventType] || `JML Event - ${employeeName}`;
  }

  private getDefaultBody(eventType: GraphEventType, employeeName: string): string {
    const bodies: Record<GraphEventType, string> = {
      ExitInterview: `<p>Exit interview meeting with ${employeeName}.</p><p>Please come prepared to discuss your experience and any feedback.</p>`,
      OnboardingSession: `<p>Welcome to the team, ${employeeName}!</p><p>This session will cover your onboarding activities and answer any questions.</p>`,
      TrainingSession: `<p>Training session for ${employeeName}.</p>`,
      EquipmentHandover: `<p>Equipment handover meeting with ${employeeName}.</p><p>Please bring all company equipment to this meeting.</p>`,
      AccessReview: `<p>Access review meeting for ${employeeName}.</p>`,
      TeamIntroduction: `<p>Team introduction meeting to welcome ${employeeName} to the team.</p>`,
      HRMeeting: `<p>HR meeting with ${employeeName}.</p>`,
      ITSetup: `<p>IT setup session for ${employeeName}.</p><p>We will configure your equipment and accounts.</p>`,
      SecurityBriefing: `<p>Security briefing for ${employeeName}.</p><p>This mandatory session covers information security policies.</p>`,
      PolicyReview: `<p>Policy review session with ${employeeName}.</p>`,
      ManagerMeeting: `<p>Meeting with your manager to discuss role and expectations.</p>`,
      BuddyIntroduction: `<p>Introduction meeting with your assigned buddy.</p>`,
      DepartmentOrientation: `<p>Department orientation for ${employeeName}.</p>`,
      Custom: `<p>Meeting with ${employeeName}.</p>`
    };
    return bodies[eventType] || `<p>Meeting with ${employeeName}.</p>`;
  }

  private buildExitInterviewBody(request: IExitInterviewRequest): string {
    return `
      <h2>Exit Interview</h2>
      <p><strong>Employee:</strong> ${request.employeeName}</p>
      <p><strong>Last Working Day:</strong> ${request.lastWorkingDay.toLocaleDateString()}</p>
      ${request.notes ? `<p><strong>Notes:</strong> ${request.notes}</p>` : ''}
      <h3>Agenda</h3>
      <ul>
        <li>Review of employment experience</li>
        <li>Feedback on team and management</li>
        <li>Suggestions for improvement</li>
        <li>Handover and transition planning</li>
        <li>Administrative matters</li>
      </ul>
      <p><em>This meeting is confidential.</em></p>
    `;
  }

  // ============================================================================
  // SharePoint Storage Operations
  // ============================================================================

  private async trackCalendarEvent(event: {
    processId: number;
    eventType: GraphEventType;
    graphEventId: string;
    scheduledDate: Date;
    attendees: string;
    status: string;
  }): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.CALENDAR_EVENTS_LIST).items.add({
        Title: `${event.eventType} - Process ${event.processId}`,
        ProcessId: event.processId,
        EventType: event.eventType,
        GraphEventId: event.graphEventId,
        ScheduledDate: event.scheduledDate,
        Attendees: event.attendees,
        Status: event.status
      });
    } catch (error) {
      logger.warn('CalendarService', 'Could not track calendar event:', error);
    }
  }

  private async getDefaultOnboardingTemplate(department?: string): Promise<IOnboardingTemplate | undefined> {
    try {
      const filter = department
        ? `(Department eq '${department}' or IsDefault eq 1) and IsActive eq 1`
        : 'IsDefault eq 1 and IsActive eq 1';

      const items = await this.sp.web.lists.getByTitle(this.ONBOARDING_TEMPLATES_LIST).items
        .filter(filter)
        .orderBy('Department', true) // Department-specific first
        .top(1)();

      if (items.length === 0) return undefined;

      const item = items[0];
      return {
        id: item.Id,
        name: item.Title,
        department: item.Department,
        role: item.Role,
        events: JSON.parse(item.Events || '[]'),
        isDefault: item.IsDefault,
        isActive: item.IsActive
      };
    } catch (error) {
      logger.warn('CalendarService', 'Could not get onboarding template:', error);
      return undefined;
    }
  }

  private async saveOnboardingSchedule(schedule: IOnboardingSchedule): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_OnboardingSchedules').items.add({
        Title: `Onboarding - ${schedule.employeeName}`,
        ProcessId: schedule.processId,
        EmployeeId: schedule.employeeId,
        EmployeeName: schedule.employeeName,
        StartDate: schedule.startDate,
        Events: JSON.stringify(schedule.events),
        Status: schedule.status
      });
    } catch (error) {
      logger.warn('CalendarService', 'Could not save onboarding schedule:', error);
    }
  }
}
