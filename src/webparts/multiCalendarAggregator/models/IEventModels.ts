import { CalendarSourceType } from './ICalendarModels';

/**
 * Extended event interface with additional metadata
 */
export interface IExtendedCalendarEvent {
  id: string;
  title: string;
  description: string;
  start: Date;
  end: Date;
  location?: string;
  category?: string;
  isAllDay: boolean;
  isRecurring: boolean;
  calendarId: string;
  calendarTitle: string;
  calendarType: CalendarSourceType;
  organizer: string;
  created: Date;
  modified: Date;
  webUrl: string;
  color: string;
  importance?: string;
  sensitivity?: string;
  showAs?: string;
  attendees?: IEventAttendee[];
  attachments?: IEventAttachment[];
  
  // Extended properties
  timeZone?: string;
  reminderMinutes?: number[];
  isPrivate?: boolean;
  isCancelled?: boolean;
  responseStatus?: string;
  meetingType?: string;
  onlineMeetingUrl?: string;
  recurrencePattern?: IRecurrencePattern;
  exceptions?: IRecurrenceException[];
  masterSeriesId?: string;
  isException?: boolean;
  tags?: string[];
  customProperties?: { [key: string]: unknown };
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event attendee interface
 */
export interface IEventAttendee {
  name: string;
  email: string;
  response: 'accepted' | 'declined' | 'tentative' | 'none';
  type: 'required' | 'optional' | 'resource';
  isOrganizer?: boolean;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event attachment interface
 */
export interface IEventAttachment {
  id?: string;
  name: string;
  contentType: string;
  size: number;
  url?: string;
  content?: string;
  isInline: boolean;
  lastModified?: Date;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Recurrence pattern interface
 */
export interface IRecurrencePattern {
  type: 'daily' | 'weekly' | 'monthly' | 'yearly' | 'weekdays' | 'custom';
  interval: number;
  daysOfWeek?: number[]; // 0 = Sunday, 1 = Monday, etc.
  dayOfMonth?: number;
  weekOfMonth?: number;
  monthOfYear?: number;
  endDate?: Date;
  occurrences?: number;
  firstDayOfWeek?: number;
  // Add index signature for ValidationUtils compatibility
  [key: string]: unknown;
}

/**
 * Recurrence exception interface
 */
export interface IRecurrenceException {
  originalDate: Date;
  newEvent?: IExtendedCalendarEvent;
  isDeleted: boolean;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event reminder interface
 */
export interface IEventReminder {
  minutes: number;
  method: 'popup' | 'email' | 'sms';
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event conflict interface
 */
export interface IEventConflict {
  event1: IExtendedCalendarEvent;
  event2: IExtendedCalendarEvent;
  conflictType: 'overlap' | 'duplicate' | 'doubleBooked';
  overlapDuration?: number;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event suggestion interface
 */
export interface IEventSuggestion {
  id: string;
  title: string;
  suggestedStart: Date;
  suggestedEnd: Date;
  reason: string;
  confidence: number;
  sourceEvent?: IExtendedCalendarEvent;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event template interface
 */
export interface IEventTemplate {
  id: string;
  name: string;
  title: string;
  description?: string;
  duration: number; // minutes
  location?: string;
  category?: string;
  attendees?: string[];
  reminders?: IEventReminder[];
  isDefault?: boolean;
  tags?: string[];
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event validation result interface
 */
export interface IEventValidationResult {
  isValid: boolean;
  errors: IEventValidationError[];
  warnings: IEventValidationWarning[];
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event validation error interface
 */
export interface IEventValidationError {
  field: string;
  message: string;
  code: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event validation warning interface
 */
export interface IEventValidationWarning {
  field: string;
  message: string;
  code: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event creation request interface
 */
export interface IEventCreateRequest {
  title: string;
  description?: string;
  start: Date;
  end: Date;
  location?: string;
  category?: string;
  isAllDay?: boolean;
  calendarId: string;
  attendees?: IEventAttendee[];
  reminders?: IEventReminder[];
  recurrence?: IRecurrencePattern;
  importance?: string;
  sensitivity?: string;
  showAs?: string;
  // Add index signature for ESLint compatibility
  [key: string]: unknown;
}

/**
 * Event update request interface
 */
export interface IEventUpdateRequest extends Partial<IEventCreateRequest> {
  id: string;
  updateRecurringSeries?: boolean;
  // Add index signature for ESLint compatibility
  [key: string]: unknown;
}

/**
 * Event search criteria interface
 */
export interface IEventSearchCriteria {
  query?: string;
  calendarIds?: string[];
  startDate?: Date;
  endDate?: Date;
  categories?: string[];
  organizers?: string[];
  attendees?: string[];
  importance?: string[];
  sensitivity?: string[];
  hasAttachments?: boolean;
  isRecurring?: boolean;
  isAllDay?: boolean;
  tags?: string[];
  customFilters?: { [key: string]: unknown };
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event search result interface
 */
export interface IEventSearchResult {
  events: IExtendedCalendarEvent[];
  totalCount: number;
  hasMore: boolean;
  nextPageToken?: string;
  searchDuration: number;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event statistics interface
 */
export interface IEventStatistics {
  totalEvents: number;
  eventsThisWeek: number;
  eventsThisMonth: number;
  upcomingEvents: number;
  pastEvents: number;
  allDayEvents: number;
  recurringEvents: number;
  eventsWithAttendees: number;
  eventsWithAttachments: number;
  eventsByCalendar: { [calendarId: string]: number };
  eventsByCategory: { [category: string]: number };
  eventsByImportance: { [importance: string]: number };
  averageEventDuration: number;
  busiestDayOfWeek: string;
  busiestHourOfDay: number;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event export options interface
 */
export interface IEventExportOptions {
  format: 'ics' | 'csv' | 'json' | 'excel';
  includeAttachments?: boolean;
  includeAttendees?: boolean;
  includeRecurrence?: boolean;
  dateRange?: {
    start: Date;
    end: Date;
  };
  calendarIds?: string[];
  categories?: string[];
  filename?: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event import result interface
 */
export interface IEventImportResult {
  successCount: number;
  errorCount: number;
  warningCount: number;
  errors: IEventImportError[];
  warnings: IEventImportWarning[];
  importedEvents: IExtendedCalendarEvent[];
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event import error interface
 */
export interface IEventImportError {
  line?: number;
  field?: string;
  message: string;
  data?: unknown;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event import warning interface
 */
export interface IEventImportWarning {
  line?: number;
  field?: string;
  message: string;
  data?: unknown;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Free/busy time slot interface
 */
export interface IFreeBusyTimeSlot {
  start: Date;
  end: Date;
  status: 'free' | 'busy' | 'tentative' | 'outOfOffice' | 'workingElsewhere';
  event?: IExtendedCalendarEvent;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Free/busy query interface
 */
export interface IFreeBusyQuery {
  attendees: string[];
  startDate: Date;
  endDate: Date;
  intervalMinutes?: number;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Free/busy result interface
 */
export interface IFreeBusyResult {
  attendee: string;
  timeSlots: IFreeBusyTimeSlot[];
  workingHours?: IWorkingHours;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Working hours interface
 */
export interface IWorkingHours {
  timeZone: string;
  daysOfWeek: number[];
  startTime: string; // HH:mm format
  endTime: string;   // HH:mm format
  // Add index signature
  [key: string]: unknown;
}

/**
 * Meeting room interface
 */
export interface IMeetingRoom {
  id: string;
  name: string;
  email: string;
  capacity: number;
  equipment?: string[];
  location?: string;
  isAvailable?: boolean;
  building?: string;
  floor?: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Meeting suggestion interface
 */
export interface IMeetingSuggestion {
  start: Date;
  end: Date;
  confidence: number;
  attendeeAvailability: { [email: string]: boolean };
  suggestedRooms?: IMeetingRoom[];
  reason: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Event notification interface
 */
export interface IEventNotification {
  id: string;
  eventId: string;
  type: 'created' | 'updated' | 'deleted' | 'reminder' | 'response';
  title: string;
  message: string;
  timestamp: Date;
  isRead: boolean;
  actionUrl?: string;
  // Add index signature
  [key: string]: unknown;
}

/**
 * Calendar synchronization status interface
 */
export interface ICalendarSyncStatus {
  calendarId: string;
  lastSyncTime: Date;
  syncStatus: 'success' | 'error' | 'inProgress' | 'pending';
  errorMessage?: string;
  nextSyncTime?: Date;
  eventCount: number;
  changes: {
    added: number;
    updated: number;
    deleted: number;
  };
  // Add index signature
  [key: string]: unknown;
}