import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Calendar source types
 */
export enum CalendarSourceType {
  SharePoint = 'SharePoint',
  SharePointList = 'SharePointList',
  SharePointCommunicationSite = 'SharePointCommunicationSite',
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
 * SharePoint List source interface
 */
export interface ISharePointListSource extends ICalendarSource {
  type: CalendarSourceType.SharePointList;
  listType: 'Events' | 'Calendar' | 'Custom' | 'Tasks' | 'Issues' | 'Announcements';
  listTemplate: number;
  viewUrl?: string;
  defaultView?: string;
  customFields?: ISharePointCustomField[];
  permissions: ISharePointPermissions;
  workflowEnabled?: boolean;
  versioningEnabled?: boolean;
  contentTypesEnabled?: boolean;
  fieldMappings?: IFieldMapping;
}

/**
 * SharePoint Communication Site source interface
 */
export interface ISharePointCommunicationSiteSource extends ICalendarSource {
  type: CalendarSourceType.SharePointCommunicationSite;
  hubSiteId?: string;
  associatedHubSites?: string[];
  newsEvents?: boolean;
  pageEvents?: boolean;
  eventWebParts?: string[];
  announcementEvents?: boolean;
  documentEvents?: boolean;
  siteActivityEvents?: boolean;
}

/**
 * SharePoint custom field interface
 */
export interface ISharePointCustomField {
  internalName: string;
  displayName: string;
  fieldType: 'Text' | 'DateTime' | 'Choice' | 'User' | 'Lookup' | 'Number' | 'Boolean' | 'URL' | 'MultiChoice' | 'Note';
  required: boolean;
  mappedTo?: 'title' | 'description' | 'location' | 'category' | 'organizer' | 'attendees' | 'custom';
  choices?: string[];
  defaultValue?: string;
  isMultiValue?: boolean;
}

/**
 * SharePoint permissions interface
 */
export interface ISharePointPermissions {
  canRead: boolean;
  canWrite: boolean;
  canDelete: boolean;
  canManagePermissions: boolean;
  canManageViews: boolean;
  canApprove?: boolean;
  canManageWebParts?: boolean;
  effectivePermissions?: string[];
  permissionLevel?: string;
}

/**
 * Field mapping interface
 */
export interface IFieldMapping {
  titleField?: string;
  startDateField: string;
  endDateField?: string;
  descriptionField?: string;
  locationField?: string;
  categoryField?: string;
  organizerField?: string;
  allDayField?: string;
  recurrenceField?: string;
  importanceField?: string;
  statusField?: string;
}

/**
 * SharePoint List configuration interface
 */
export interface ISharePointListConfiguration {
  includeCustomLists: boolean;
  includeTaskLists: boolean;
  includeIssueLists: boolean;
  includeAnnouncementLists: boolean;
  autoDiscoverDateFields: boolean;
  enableWorkflowIntegration: boolean;
  enableVersionHistory: boolean;
  enableComments: boolean;
  maxItemsPerList: number;
  dateRangeMonths: number;
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
 * SharePoint event interface
 */
export interface ISharePointEvent extends ICalendarEvent {
  listItemId?: number;
  listItemUrl?: string;
  workflowStatus?: string;
  approvalStatus?: 'Approved' | 'Pending' | 'Rejected' | 'Draft';
  customFields?: { [fieldName: string]: unknown };
  sharePointAttachments?: ISharePointAttachment[];
  versions?: ISharePointVersion[];
  comments?: ISharePointComment[];
  contentType?: string;
  etag?: string;
  hasAttachments?: boolean;
  percentComplete?: number;
  assignedTo?: string[];
  priority?: string;
  taskStatus?: string;
  newsCategory?: string;
  pageLayout?: string;
  publishedDate?: Date;
  expirationDate?: Date;
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
  [key: string]: unknown;
}

/**
 * SharePoint attachment interface
 */
export interface ISharePointAttachment {
  fileName: string;
  serverRelativeUrl: string;
  size: number;
  created: Date;
  createdBy: string;
  lastModified: Date;
  contentType?: string;
  checksum?: string;
}

/**
 * SharePoint version interface
 */
export interface ISharePointVersion {
  versionId: number;
  versionLabel: string;
  created: Date;
  createdBy: string;
  changeDescription?: string;
  size?: number;
  isCurrentVersion?: boolean;
}

/**
 * SharePoint comment interface
 */
export interface ISharePointComment {
  id: number;
  text: string;
  author: string;
  authorEmail?: string;
  created: Date;
  modified?: Date;
  replies?: ISharePointComment[];
  isDeleted?: boolean;
  likeCount?: number;
  mentionedUsers?: string[];
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
    includeCustomLists: boolean;
    includeCommunicationSites: boolean;
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

/**
 * SharePoint Communication Site configuration interface
 */
export interface ISharePointCommunicationSiteConfiguration {
  includeNewsEvents: boolean;
  includePageEvents: boolean;
  includeAnnouncementEvents: boolean;
  includeDocumentEvents: boolean;
  includeSiteActivityEvents: boolean;
  newsDateRange: number; // days
  pageUpdateEvents: boolean;
  documentUploadEvents: boolean;
  maxEventsPerSite: number;
}

/**
 * Web part instance data interface
 */
export interface IWebPartInstanceData {
  id: string;
  webPartType: string;
  title: string;
  properties: Record<string, unknown>;
  pageId: string;
  pageTitle: string;
  zoneIndex: number;
  order: number;
  isVisible: boolean;
}

/**
 * News post interface
 */
export interface INewsPost {
  id: number;
  title: string;
  description?: string;
  authorByline?: string;
  bannerImageUrl?: string;
  created: Date;
  modified: Date;
  publishedDate?: Date;
  firstPublishedDate?: Date;
  canvasContent1?: string;
  promotedState?: number;
  topicHeader?: string;
  url: string;
  authorId?: number;
  authorDisplayName?: string;
}

/**
 * Site activity interface
 */
export interface ISiteActivity {
  id: string;
  activityType: 'PageCreated' | 'PageModified' | 'NewsPublished' | 'DocumentUploaded' | 'ListItemCreated' | 'ListItemModified';
  title: string;
  description?: string;
  actor: string;
  actorEmail?: string;
  timestamp: Date;
  resourceUrl: string;
  resourceTitle: string;
  siteId: string;
  webId: string;
  listId?: string;
  itemId?: number;
}