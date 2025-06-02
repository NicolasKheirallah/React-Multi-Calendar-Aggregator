export class AppConstants {
  // Application metadata
  public static readonly APP_NAME = 'Multi-Calendar Aggregator';
  public static readonly APP_VERSION = '1.0.0';
  public static readonly APP_DESCRIPTION = 'Unified calendar view from multiple SharePoint and Exchange sources';

  // SharePoint constants
  public static readonly SHAREPOINT_CALENDAR_BASE_TEMPLATE = 106;
  public static readonly SHAREPOINT_LIST_ITEM_LIMIT = 5000;
  public static readonly SHAREPOINT_BATCH_SIZE = 100;

  // Microsoft Graph constants
  public static readonly GRAPH_API_VERSION = 'v1.0';
  public static readonly GRAPH_CALENDARS_ENDPOINT = '/me/calendars';
  public static readonly GRAPH_EVENTS_ENDPOINT = '/me/events';
  public static readonly GRAPH_CALENDAR_EVENTS_ENDPOINT = '/me/calendars/{calendarId}/events';

  // Cache constants
  public static readonly CACHE_KEY_PREFIX = 'multi-calendar-aggregator-';
  public static readonly CACHE_DURATION_MINUTES = 15;
  public static readonly CACHE_SOURCES_KEY = 'calendar-sources';
  public static readonly CACHE_EVENTS_KEY = 'calendar-events';

  // UI constants
  public static readonly DEFAULT_VIEW_TYPE = 'month';
  public static readonly DEFAULT_MAX_EVENTS = 100;
  public static readonly DEFAULT_REFRESH_INTERVAL = 5; // minutes
  public static readonly DEFAULT_DATE_FORMAT = 'MMMM Do, YYYY';
  public static readonly DEFAULT_TIME_FORMAT = 'h:mm A';
  public static readonly DEFAULT_DATETIME_FORMAT = 'MMMM Do, YYYY [at] h:mm A';

  // Event display constants
  public static readonly EVENT_TITLE_MAX_LENGTH = 50;
  public static readonly EVENT_DESCRIPTION_MAX_LENGTH = 200;
  public static readonly EVENT_LOCATION_MAX_LENGTH = 100;

  // Calendar view constants
  public static readonly CALENDAR_VIEWS = {
    MONTH: 'month',
    WEEK: 'week',
    DAY: 'day',
    AGENDA: 'agenda',
    TIMELINE: 'timeline'
  } as const;

  // Time constants
  public static readonly MILLISECONDS_IN_MINUTE = 60 * 1000;
  public static readonly MILLISECONDS_IN_HOUR = 60 * 60 * 1000;
  public static readonly MILLISECONDS_IN_DAY = 24 * 60 * 60 * 1000;
  public static readonly MILLISECONDS_IN_WEEK = 7 * 24 * 60 * 60 * 1000;

  // Recurrence patterns
  public static readonly RECURRENCE_PATTERNS = {
    DAILY: 'daily',
    WEEKLY: 'weekly',
    MONTHLY: 'monthly',
    YEARLY: 'yearly',
    WEEKDAYS: 'weekdays',
    CUSTOM: 'custom'
  } as const;

  // Event importance levels
  public static readonly EVENT_IMPORTANCE = {
    LOW: 'low',
    NORMAL: 'normal',
    HIGH: 'high'
  } as const;

  // Event sensitivity levels
  public static readonly EVENT_SENSITIVITY = {
    NORMAL: 'normal',
    PERSONAL: 'personal',
    PRIVATE: 'private',
    CONFIDENTIAL: 'confidential'
  } as const;

  // Response status
  public static readonly RESPONSE_STATUS = {
    NONE: 'none',
    ACCEPTED: 'accepted',
    DECLINED: 'declined',
    TENTATIVE: 'tentative'
  } as const;

  // Show as status
  public static readonly SHOW_AS_STATUS = {
    FREE: 'free',
    TENTATIVE: 'tentative',
    BUSY: 'busy',
    OOF: 'oof',
    WORKING_ELSEWHERE: 'workingElsewhere'
  } as const;

  // Error messages
  public static readonly ERROR_MESSAGES = {
    GENERAL_ERROR: 'An unexpected error occurred. Please try again.',
    PERMISSION_DENIED: 'You do not have permission to access this calendar.',
    NETWORK_ERROR: 'Network error. Please check your connection.',
    INVALID_CONFIGURATION: 'Invalid configuration. Please check your settings.',
    CALENDAR_NOT_FOUND: 'Calendar not found or no longer accessible.',
    EVENTS_LOAD_FAILED: 'Failed to load events from one or more calendars.',
    GRAPH_API_ERROR: 'Microsoft Graph API error. Please try again later.',
    SHAREPOINT_API_ERROR: 'SharePoint API error. Please check your permissions.',
    TIMEOUT_ERROR: 'Request timed out. Please try again.',
    QUOTA_EXCEEDED: 'API quota exceeded. Please try again later.'
  } as const;

  // Success messages
  public static readonly SUCCESS_MESSAGES = {
    CALENDARS_LOADED: 'Calendars loaded successfully.',
    EVENTS_REFRESHED: 'Events refreshed successfully.',
    SETTINGS_SAVED: 'Settings saved successfully.',
    EXPORT_COMPLETED: 'Calendar export completed.',
    CALENDAR_ADDED: 'Calendar source added successfully.',
    CALENDAR_REMOVED: 'Calendar source removed successfully.'
  } as const;

  // API endpoints and parameters
  public static readonly API_LIMITS = {
    SHAREPOINT_LIST_THRESHOLD: 5000,
    GRAPH_REQUEST_TIMEOUT: 30000, // 30 seconds
    MAX_EVENTS_PER_REQUEST: 999,
    MAX_CALENDARS_PER_USER: 100,
    MAX_CONCURRENT_REQUESTS: 10
  } as const;

  // SharePoint REST API
  public static readonly SHAREPOINT_API = {
    LISTS_ENDPOINT: '/_api/web/lists',
    CALENDAR_FILTER: "BaseTemplate eq 106",
    LIST_ITEMS_ENDPOINT: '/_api/web/lists(guid\'{listId}\')/items',
    SITES_ENDPOINT: '/_api/web/webs',
    CURRENT_USER_ENDPOINT: '/_api/web/currentuser'
  } as const;

  // Microsoft Graph scopes
  public static readonly GRAPH_SCOPES = {
    CALENDARS_READ: 'Calendars.Read',
    CALENDARS_READ_SHARED: 'Calendars.Read.Shared',
    USER_READ: 'User.Read',
    GROUP_READ_ALL: 'Group.Read.All'
  } as const;

  // File export formats
  public static readonly EXPORT_FORMATS = {
    ICS: 'ics',
    CSV: 'csv',
    JSON: 'json',
    XLSX: 'xlsx'
  } as const;

  // Local storage keys
  public static readonly STORAGE_KEYS = {
    USER_PREFERENCES: 'multi-cal-user-prefs',
    LAST_REFRESH: 'multi-cal-last-refresh',
    SELECTED_CALENDARS: 'multi-cal-selected',
    VIEW_SETTINGS: 'multi-cal-view-settings',
    FILTER_SETTINGS: 'multi-cal-filters'
  } as const;

  // Event categories (common SharePoint/Exchange categories)
  public static readonly EVENT_CATEGORIES = [
    'Meeting',
    'Appointment',
    'Conference',
    'Personal',
    'Holiday',
    'Training',
    'Deadline',
    'Review',
    'Travel',
    'Other'
  ] as const;

  // Time zones (common ones)
  public static readonly COMMON_TIMEZONES = [
    'UTC',
    'America/New_York',
    'America/Chicago',
    'America/Denver',
    'America/Los_Angeles',
    'Europe/London',
    'Europe/Paris',
    'Europe/Berlin',
    'Asia/Tokyo',
    'Asia/Shanghai',
    'Australia/Sydney'
  ] as const;

  // Validation rules
  public static readonly VALIDATION = {
    MIN_REFRESH_INTERVAL: 1, // minutes
    MAX_REFRESH_INTERVAL: 60, // minutes
    MIN_MAX_EVENTS: 10,
    MAX_MAX_EVENTS: 1000,
    MIN_TITLE_LENGTH: 1,
    MAX_TITLE_LENGTH: 100
  } as const;

  // Feature flags
  public static readonly FEATURES = {
    ENABLE_GRAPH_API: true,
    ENABLE_CACHING: true,
    ENABLE_EXPORT: true,
    ENABLE_NOTIFICATIONS: true,
    ENABLE_SEARCH: true,
    ENABLE_FILTERS: true,
    ENABLE_SHARING: true,
    ENABLE_ANALYTICS: false
  } as const;

  // Performance constants
  public static readonly PERFORMANCE = {
    DEBOUNCE_DELAY: 300, // milliseconds
    ANIMATION_DURATION: 200, // milliseconds
    VIRTUAL_SCROLL_BUFFER: 50, // items
    IMAGE_LAZY_LOAD_THRESHOLD: 100, // pixels
    REQUEST_RETRY_COUNT: 3,
    REQUEST_RETRY_DELAY: 1000 // milliseconds
  } as const;

  // SharePoint field names
  public static readonly SHAREPOINT_FIELDS = {
    TITLE: 'Title',
    DESCRIPTION: 'Description',
    EVENT_DATE: 'EventDate',
    END_DATE: 'EndDate',
    LOCATION: 'Location',
    CATEGORY: 'Category',
    ALL_DAY_EVENT: 'fAllDayEvent',
    RECURRENCE: 'fRecurrence',
    RECURRENCE_DATA: 'RecurrenceData',
    AUTHOR: 'Author',
    EDITOR: 'Editor',
    CREATED: 'Created',
    MODIFIED: 'Modified',
    ID: 'Id'
  } as const;

  // Graph API field selections
  public static readonly GRAPH_EVENT_FIELDS = [
    'id',
    'subject',
    'body',
    'start',
    'end',
    'location',
    'categories',
    'isAllDay',
    'recurrence',
    'organizer',
    'attendees',
    'createdDateTime',
    'lastModifiedDateTime',
    'webLink',
    'importance',
    'sensitivity',
    'showAs'
  ].join(',');

  public static readonly GRAPH_CALENDAR_FIELDS = [
    'id',
    'name',
    'color',
    'canEdit',
    'canShare',
    'canViewPrivateItems',
    'owner'
  ].join(',');

  // CSS class names
  public static readonly CSS_CLASSES = {
    CONTAINER: 'multi-calendar-container',
    HEADER: 'multi-calendar-header',
    CALENDAR_VIEW: 'calendar-view',
    EVENT_CARD: 'event-card',
    LOADING: 'loading-state',
    ERROR: 'error-state',
    AGENDA_VIEW: 'agenda-view',
    TIMELINE_VIEW: 'timeline-view'
  } as const;

  // Accessibility constants
  public static readonly ACCESSIBILITY = {
    ARIA_LABELS: {
      CALENDAR: 'Calendar view',
      EVENT: 'Calendar event',
      NAVIGATION: 'Calendar navigation',
      FILTERS: 'Calendar filters',
      SEARCH: 'Search events'
    },
    KEYBOARD_SHORTCUTS: {
      NEXT: 'ArrowRight',
      PREVIOUS: 'ArrowLeft',
      TODAY: 't',
      SEARCH: '/',
      ESCAPE: 'Escape'
    }
  } as const;

  // Regular expressions
  public static readonly REGEX = {
    EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    HEX_COLOR: /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/,
    GUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i,
    TIME_24H: /^([01]\d|2[0-3]):([0-5]\d)$/
  } as const;

  // Event icons mapping
  public static readonly EVENT_ICONS = {
    MEETING: 'People',
    APPOINTMENT: 'Calendar',
    CONFERENCE: 'Video',
    PERSONAL: 'Contact',
    HOLIDAY: 'Sunny',
    TRAINING: 'Education',
    DEADLINE: 'Clock',
    REVIEW: 'ReviewSolid',
    TRAVEL: 'Airplane',
    OTHER: 'More'
  } as const;

  // Color themes
  public static readonly THEMES = {
    DEFAULT: 'default',
    BLUE: 'blue',
    GREEN: 'green',
    ORANGE: 'orange',
    PURPLE: 'purple',
    RED: 'red'
  } as const;

  // Application URLs
  public static readonly URLS = {
    DOCUMENTATION: 'https://docs.microsoft.com/sharepoint/dev/spfx/',
    SUPPORT: 'mailto:support@yourcompany.com',
    PRIVACY: 'https://privacy.microsoft.com/',
    TERMS: 'https://www.microsoft.com/servicesagreement/'
  } as const;
}