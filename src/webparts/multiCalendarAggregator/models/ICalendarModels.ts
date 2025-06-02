import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Calendar source types
 */
export enum CalendarSourceType {
  SharePoint = 'SharePoint',
  Exchange = 'Exchange'
}

/**
 * Calendar source interface
 */
export interface ICalendarSource {
  id: string;
  title: string;
  description: string;
  type: CalendarSourceType;
  url: string;
  siteTitle: string;
  siteUrl: string;
  color: string;
  isEnabled: boolean;
  itemCount?: number;
  lastModified?: string;
  canEdit?: boolean;
  canShare?: boolean;
}

/**
 * Calendar event interface
 */
export interface ICalendarEvent {
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
  tags?: string[];
  masterSeriesId?: string;
  isException?: boolean;
}

/**
 * Event attendee interface
 */
export interface IEventAttendee {
  name: string;
  email: string;
  response: 'accepted' | 'declined' | 'tentative' | 'none';
  type: 'required' | 'optional' | 'resource';
  isOrganizer?: boolean; // Add this field to match IEventModels
  [key: string]: unknown;
}

/**
 * Event attachment interface
 */
export interface IEventAttachment {
  id?: string; // Add optional id field
  name: string;
  contentType: string;
  size: number;
  url?: string;
  content?: string; // Add optional content field
  isInline: boolean;
  lastModified?: Date; // Add optional lastModified field
  [key: string]: unknown;
}


/**
 * Calendar view types
 */
export type CalendarViewType = 'month' | 'week' | 'day' | 'agenda' | 'timeline';

/**
 * Event filter options
 */
export interface IEventFilter {
  calendarIds?: string[];
  categories?: string[];
  startDate?: Date;
  endDate?: Date;
  searchQuery?: string;
  showAllDay?: boolean;
  showRecurring?: boolean;
  importance?: string[];
}

/**
 * Calendar statistics
 */
export interface ICalendarStats {
  totalEvents: number;
  upcomingEvents: number;
  todaysEvents: number;
  thisWeekEvents: number;
  thisMonthEvents: number;
  eventsByCalendar: { [calendarId: string]: number };
  eventsByCategory: { [category: string]: number };
}

/**
 * Calendar service interface
 */
export interface ICalendarService {
  initialize(): Promise<void>;
  getCalendarSources(includeExchange?: boolean): Promise<ICalendarSource[]>;
  getEventsFromSource(source: ICalendarSource, maxEvents?: number): Promise<ICalendarEvent[]>;
  getEventsFromSources(sources: ICalendarSource[], maxEvents?: number): Promise<ICalendarEvent[]>;
  searchEvents(sources: ICalendarSource[], query: string, maxResults?: number): Promise<ICalendarEvent[]>;
  clearCache(): void;
}

/**
 * Multi-calendar aggregator props
 */
export interface IMultiCalendarAggregatorProps {
  title: string;
  selectedCalendars: string[];
  viewType: CalendarViewType;
  showWeekends: boolean;
  maxEvents: number;
  refreshInterval: number;
  useGraphAPI: boolean;
  colorCoding: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

/**
 * Event details panel props
 */
export interface IEventDetailsPanelProps {
  event: ICalendarEvent;
  calendarSource?: ICalendarSource;
  onClose: () => void;
  onEdit?: (event: ICalendarEvent) => void;
  onDelete?: (eventId: string) => void;
}

/**
 * Calendar sources panel props
 */
export interface ICalendarSourcesPanelProps {
  sources: ICalendarSource[];
  selectedSources: string[];
  onSourcesChange: (selectedIds: string[]) => void;
  onClose: () => void;
  onRefresh?: () => void;
}

/**
 * Agenda view props
 */
export interface IAgendaViewProps {
  events: ICalendarEvent[];
  onEventSelect: (event: ICalendarEvent) => void;
  calendarSources: ICalendarSource[];
  theme: unknown;
  groupBy?: 'date' | 'calendar' | 'category';
  showDays?: number;
}

/**
 * Timeline view props
 */
export interface ITimelineViewProps {
  events: ICalendarEvent[];
  onEventSelect: (event: ICalendarEvent) => void;
  calendarSources: ICalendarSource[];
  theme: unknown;
  currentDate: Date;
  timeRange?: 'day' | 'week' | 'month';
}

/**
 * Calendar theme configuration
 */
export interface ICalendarTheme {
  primary: string;
  primaryDark: string;
  primaryLight: string;
  secondary: string;
  background: string;
  surface: string;
  text: string;
  textSecondary: string;
  border: string;
  shadow: string;
  success: string;
  warning: string;
  error: string;
  info: string;
}

/**
 * Export/import formats
 */
export enum ExportFormat {
  ICS = 'ics',
  CSV = 'csv',
  JSON = 'json',
  XLSX = 'xlsx'
}

/**
 * Calendar permissions
 */
export interface ICalendarPermissions {
  canRead: boolean;
  canWrite: boolean;
  canDelete: boolean;
  canShare: boolean;
  canManagePermissions: boolean;
}

/**
 * Notification settings
 */
export interface INotificationSettings {
  enabled: boolean;
  emailNotifications: boolean;
  browserNotifications: boolean;
  reminderMinutes: number[];
  digestFrequency: 'none' | 'daily' | 'weekly';
}

/**
 * Calendar sync status
 */
export interface ISyncStatus {
  calendarId: string;
  lastSync: Date;
  status: 'success' | 'error' | 'syncing';
  errorMessage?: string;
  nextSync?: Date;
}

/**
 * Advanced filter options
 */
export interface IAdvancedFilter {
  dateRange: {
    start: Date;
    end: Date;
    preset?: 'today' | 'tomorrow' | 'this-week' | 'next-week' | 'this-month' | 'next-month' | 'custom';
  };
  timeRange?: {
    start: string; // HH:mm format
    end: string;   // HH:mm format
  };
  daysOfWeek: number[]; // 0-6, Sunday = 0
  eventTypes: {
    meetings: boolean;
    appointments: boolean;
    allDay: boolean;
    recurring: boolean;
    private: boolean;
  };
  priorities: string[];
  attendeeFilter?: {
    includeOrganized: boolean;
    includeAttending: boolean;
    excludeDeclined: boolean;
  };
}

/**
 * Calendar integration settings
 */
export interface IIntegrationSettings {
  sharePoint: {
    enabled: boolean;
    includeSites: string[];
    excludeSites: string[];
    includeSubsites: boolean;
  };
  exchange: {
    enabled: boolean;
    includeSharedCalendars: boolean;
    includeRoomCalendars: boolean;
    includeResourceCalendars: boolean;
  };
  teams: {
    enabled: boolean;
    includeChannelMeetings: boolean;
    includePrivateMeetings: boolean;
  };
  outlook: {
    enabled: boolean;
    syncCategories: boolean;
    syncReminders: boolean;
  };
}