export interface ICalendarEvent {
  id: number;
  title: string;
  description: string;
  startDate: Date;
  endDate: Date;
  isAllDay: boolean;

  // Process/Event Type
  processType: ProcessType;
  processInstanceId?: number; // Link to specific process instance

  // Location & Attendees
  location?: string;
  attendees?: string[]; // Array of email addresses
  organizer?: string;

  // Status & Tracking
  status: EventStatus;
  isRecurring: boolean;
  recurrenceRule?: string; // iCal RRULE format

  // Colors & Styling
  color?: string;
  backgroundColor?: string;

  // Milestone & Dependencies
  isMilestone: boolean;
  dependsOn?: number[]; // Array of event IDs

  // Outlook Integration
  outlookEventId?: string;
  isSyncedWithOutlook: boolean;

  // Metadata
  createdBy: string;
  createdDate: Date;
  modifiedBy: string;
  modifiedDate: Date;
}

export enum ProcessType {
  Onboarding = 'Onboarding',
  Offboarding = 'Offboarding',
  Probation = 'Probation',
  Training = 'Training',
  Review = 'Review',
  Interview = 'Interview',
  Meeting = 'Meeting',
  Deadline = 'Deadline',
  Other = 'Other'
}

export enum EventStatus {
  Scheduled = 'Scheduled',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Rescheduled = 'Rescheduled',
  Pending = 'Pending'
}

export type CalendarView = 'month' | 'week' | 'day' | 'team';

export interface IJmlCalendarProps {
  sp: any;
  siteUrl: string;
  userEmail: string;
  userDisplayName: string;

  // View Settings
  viewMode: CalendarView;
  defaultDate: string;
  firstDayOfWeek: number;

  // Display Options
  showWeekends: boolean;
  showWeekNumbers: boolean;
  showEventCategories: boolean;
  showTimeSlots: boolean;
  compactView: boolean;

  // Event Settings
  defaultEventDuration: number;
  allowEventCreation: boolean;
  showRecurringEvents: boolean;
  maxEventsPerDay: number;

  // Filter Settings
  defaultEventType: string;
  showAllDayEvents: boolean;
  showPastEvents: boolean;

  // Performance Settings
  monthsToLoad: number;
  autoRefresh: boolean;
  refreshInterval: number;
}

export interface IJmlCalendarState {
  loading: boolean;
  error: string;
  events: ICalendarEvent[];
  currentDate: Date;
  viewMode: CalendarView;
  selectedEvent?: ICalendarEvent;
  showEventDetails: boolean;
  filterProcessType?: ProcessType;
  showTeamCalendar: boolean;
}

// For team calendar view
export interface ITeamMember {
  email: string;
  displayName: string;
  department: string;
  events: ICalendarEvent[];
}

// For drag-and-drop
export interface IDragDropContext {
  eventId: number;
  originalStartDate: Date;
  originalEndDate: Date;
}

// For Outlook Integration
export interface IOutlookSyncResult {
  success: boolean;
  syncedCount: number;
  failedCount: number;
  errors?: string[];
}

// For iCal Export
export interface ICalExportOptions {
  includeCompleted: boolean;
  dateRange?: {
    start: Date;
    end: Date;
  };
  processTypes?: ProcessType[];
}

// Calendar Day Cell
export interface ICalendarDay {
  date: Date;
  isCurrentMonth: boolean;
  isToday: boolean;
  events: ICalendarEvent[];
}

// Week View
export interface ICalendarWeek {
  weekNumber: number;
  days: ICalendarDay[];
}

// Month View
export interface ICalendarMonth {
  month: number;
  year: number;
  weeks: ICalendarWeek[];
}

// ============================================================================
// Microsoft Graph Calendar Types
// ============================================================================

export type GraphEventType =
  | 'ExitInterview'
  | 'OnboardingSession'
  | 'TrainingSession'
  | 'EquipmentHandover'
  | 'AccessReview'
  | 'TeamIntroduction'
  | 'HRMeeting'
  | 'ITSetup'
  | 'SecurityBriefing'
  | 'PolicyReview'
  | 'ManagerMeeting'
  | 'BuddyIntroduction'
  | 'DepartmentOrientation'
  | 'Custom';

export type GraphResponseStatus = 'None' | 'Organizer' | 'TentativelyAccepted' | 'Accepted' | 'Declined' | 'NotResponded';

export type MeetingType = 'InPerson' | 'Online' | 'Hybrid';

export interface IGraphCalendarEvent {
  id?: string;
  subject: string;
  bodyPreview?: string;
  body?: {
    contentType: 'Text' | 'HTML';
    content: string;
  };
  start: IGraphDateTimeZone;
  end: IGraphDateTimeZone;
  location?: IGraphLocation;
  attendees: IGraphAttendee[];
  organizer?: IGraphOrganizer;
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: 'TeamsForBusiness' | 'SkypeForBusiness' | 'Unknown';
  onlineMeeting?: IGraphOnlineMeetingInfo;
  showAs?: 'Free' | 'Tentative' | 'Busy' | 'Oof' | 'WorkingElsewhere' | 'Unknown';
  importance?: 'Low' | 'Normal' | 'High';
  sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  categories?: string[];
  isAllDay?: boolean;
  isCancelled?: boolean;
  reminderMinutesBeforeStart?: number;
  responseRequested?: boolean;
  webLink?: string;
  createdDateTime?: Date;
  lastModifiedDateTime?: Date;
}

export interface IGraphDateTimeZone {
  dateTime: string; // ISO 8601 format
  timeZone: string; // e.g., "UTC", "Pacific Standard Time"
}

export interface IGraphLocation {
  displayName: string;
  locationType?: 'Default' | 'ConferenceRoom' | 'HomeAddress' | 'BusinessAddress';
  uniqueId?: string;
  locationEmailAddress?: string;
  address?: {
    street?: string;
    city?: string;
    state?: string;
    countryOrRegion?: string;
    postalCode?: string;
  };
}

export interface IGraphAttendee {
  emailAddress: IGraphEmailAddress;
  type: 'Required' | 'Optional' | 'Resource';
  status?: {
    response: GraphResponseStatus;
    time?: Date;
  };
}

export interface IGraphEmailAddress {
  address: string;
  name?: string;
}

export interface IGraphOrganizer {
  emailAddress: IGraphEmailAddress;
}

export interface IGraphOnlineMeetingInfo {
  joinUrl?: string;
  conferenceId?: string;
  tollNumber?: string;
  tollFreeNumbers?: string[];
  dialinUrl?: string;
  quickDial?: string;
}

// ============================================================================
// Meeting Room / Resource Types
// ============================================================================

export interface IGraphMeetingRoom {
  id: string;
  emailAddress: string;
  displayName: string;
  capacity?: number;
  building?: string;
  floorNumber?: number;
  floorLabel?: string;
  isWheelChairAccessible?: boolean;
  bookingType?: 'Standard' | 'Reserved';
  tags?: string[];
  address?: {
    street?: string;
    city?: string;
    state?: string;
    postalCode?: string;
  };
}

export interface IRoomAvailability {
  room: IGraphMeetingRoom;
  availability: 'Available' | 'Busy' | 'Tentative' | 'Unknown';
  availabilityView?: string;
}

// ============================================================================
// Schedule Finding Types
// ============================================================================

export interface IScheduleRequest {
  schedules: string[]; // Email addresses
  startTime: IGraphDateTimeZone;
  endTime: IGraphDateTimeZone;
  availabilityViewInterval?: number; // Minutes (15, 30, 60)
}

export interface IScheduleResponse {
  scheduleId: string;
  availabilityView: string; // Binary string: 0=free, 1=tentative, 2=busy, 3=oof
  scheduleItems: IScheduleItem[];
  workingHours?: IWorkingHours;
  error?: {
    message: string;
    responseCode: string;
  };
}

export interface IScheduleItem {
  status: 'Free' | 'Tentative' | 'Busy' | 'Oof' | 'WorkingElsewhere' | 'Unknown';
  start: IGraphDateTimeZone;
  end: IGraphDateTimeZone;
  subject?: string;
  location?: string;
  isPrivate?: boolean;
}

export interface IWorkingHours {
  daysOfWeek: string[];
  startTime: string; // HH:mm:ss
  endTime: string;
  timeZone: {
    name: string;
  };
}

export interface IMeetingTimeSuggestion {
  confidence: number; // 0-100
  organizerAvailability: 'Free' | 'Tentative' | 'Busy' | 'Oof' | 'WorkingElsewhere' | 'Unknown';
  attendeeAvailability: Array<{
    attendee: { emailAddress: IGraphEmailAddress };
    availability: 'Free' | 'Tentative' | 'Busy' | 'Oof' | 'WorkingElsewhere' | 'Unknown';
  }>;
  locations: IGraphLocation[];
  meetingTimeSlot: {
    start: IGraphDateTimeZone;
    end: IGraphDateTimeZone;
  };
  suggestionReason?: string;
}

// ============================================================================
// JML-Specific Calendar Types
// ============================================================================

export interface IJmlGraphCalendarEvent extends IGraphCalendarEvent {
  processId?: number;
  processType?: 'Joiner' | 'Mover' | 'Leaver';
  eventType: GraphEventType;
  eventStatus: EventStatus;
  employeeId?: string;
  employeeName?: string;
  relatedTaskId?: number;
  meetingType: MeetingType;
  notes?: string;
}

export interface IEventTemplate {
  id: number;
  name: string;
  eventType: GraphEventType;
  defaultSubject: string;
  defaultBody: string;
  defaultDuration: number; // Minutes
  defaultMeetingType: MeetingType;
  isOnlineMeeting: boolean;
  defaultAttendeeRoles?: string[]; // E.g., ['HR', 'Manager', 'Employee']
  defaultReminderMinutes: number;
  processTypes: ('Joiner' | 'Mover' | 'Leaver')[];
  isActive: boolean;
}

export interface IScheduleEventRequest {
  processId: number;
  eventType: GraphEventType;
  employeeEmail: string;
  employeeName: string;
  organizerEmail: string;
  additionalAttendees?: string[];
  preferredDate?: Date;
  preferredTimeRange?: {
    startHour: number; // 0-23
    endHour: number;
  };
  duration: number; // Minutes
  meetingType: MeetingType;
  roomRequirements?: {
    capacity?: number;
    building?: string;
    isWheelChairAccessible?: boolean;
  };
  subject?: string;
  body?: string;
  isOnlineMeeting?: boolean;
}

export interface IScheduleEventResult {
  success: boolean;
  eventId?: string;
  scheduledTime?: IGraphDateTimeZone;
  meetingRoom?: IGraphMeetingRoom;
  onlineMeetingUrl?: string;
  errorMessage?: string;
  conflictDetails?: string;
  alternativeTimes?: IMeetingTimeSuggestion[];
}

// ============================================================================
// Onboarding Schedule Types
// ============================================================================

export interface IOnboardingSchedule {
  processId: number;
  employeeId: string;
  employeeName: string;
  startDate: Date;
  events: IOnboardingEvent[];
  status: 'Draft' | 'Scheduled' | 'InProgress' | 'Completed';
  createdAt: Date;
  modifiedAt: Date;
}

export interface IOnboardingEvent {
  dayOffset: number; // Days from start date (0 = first day)
  eventType: GraphEventType;
  subject: string;
  duration: number; // Minutes
  organizer: string; // Email
  attendees: string[]; // Emails
  isRequired: boolean;
  isOnlineMeeting: boolean;
  scheduledEventId?: string; // After scheduling
  status: EventStatus;
}

export interface IOnboardingTemplate {
  id: number;
  name: string;
  department?: string;
  role?: string;
  events: IOnboardingEvent[];
  isDefault: boolean;
  isActive: boolean;
}

// ============================================================================
// Exit Interview Types
// ============================================================================

export interface IExitInterviewRequest {
  processId: number;
  employeeEmail: string;
  employeeName: string;
  lastWorkingDay: Date;
  hrContactEmail: string;
  managerEmail?: string;
  preferredDate?: Date;
  duration?: number; // Default 60 minutes
  isOnlineMeeting?: boolean;
  includeManager?: boolean;
  notes?: string;
}

export interface IExitInterviewResult {
  success: boolean;
  eventId?: string;
  scheduledDateTime?: Date;
  attendees: string[];
  meetingUrl?: string;
  errorMessage?: string;
}
