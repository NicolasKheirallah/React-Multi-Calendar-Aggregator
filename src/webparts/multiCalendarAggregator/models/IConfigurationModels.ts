import { CalendarViewType } from './ICalendarModels';

/**
 * Web part configuration interface
 */
export interface IWebPartConfiguration {
  title: string;
  description?: string;
  defaultView: CalendarViewType;
  maxEvents: number;
  refreshInterval: number;
  showWeekends: boolean;
  enableGraphAPI: boolean;
  enableColorCoding: boolean;
  enableExport: boolean;
  enableSearch: boolean;
  enableFilters: boolean;
  enableNotifications: boolean;
  theme: IThemeConfiguration;
  layout: ILayoutConfiguration;
  permissions: IPermissionConfiguration;
  integration: IIntegrationConfiguration;
}

/**
 * Theme configuration interface
 */
export interface IThemeConfiguration {
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  backgroundColor: string;
  textColor: string;
  borderColor: string;
  useSystemTheme: boolean;
  customCss?: string;
  colorPalette: string[];
}

/**
 * Layout configuration interface
 */
export interface ILayoutConfiguration {
  compactMode: boolean;
  showHeader: boolean;
  showToolbar: boolean;
  showSidebar: boolean;
  sidebarPosition: 'left' | 'right';
  headerHeight: number;
  toolbarHeight: number;
  sidebarWidth: number;
  responsive: IResponsiveConfiguration;
}

/**
 * Responsive configuration interface
 */
export interface IResponsiveConfiguration {
  breakpoints: {
    mobile: number;
    tablet: number;
    desktop: number;
  };
  mobileView: CalendarViewType;
  tabletView: CalendarViewType;
  desktopView: CalendarViewType;
  hideSidebarOnMobile: boolean;
  collapseToolbarOnMobile: boolean;
}

/**
 * Permission configuration interface
 */
export interface IPermissionConfiguration {
  allowCalendarSelection: boolean;
  allowViewChange: boolean;
  allowExport: boolean;
  allowPrint: boolean;
  allowShare: boolean;
  restrictedCalendars: string[];
  hiddenCalendars: string[];
  readOnlyMode: boolean;
}

/**
 * Integration configuration interface
 */
export interface IIntegrationConfiguration {
  sharePoint: ISharePointIntegration;
  exchange: IExchangeIntegration;
  teams: ITeamsIntegration;
  outlook: IOutlookIntegration;
  powerPlatform: IPowerPlatformIntegration;
}

/**
 * SharePoint integration configuration
 */
export interface ISharePointIntegration {
  enabled: boolean;
  includeSubsites: boolean;
  includeParentSite: boolean;
  siteUrls: string[];
  excludedSites: string[];
  listTemplateIds: number[];
  customFields: ICustomFieldMapping[];
}

/**
 * Exchange integration configuration
 */
export interface IExchangeIntegration {
  enabled: boolean;
  includeSharedCalendars: boolean;
  includeGroupCalendars: boolean;
  includeRoomCalendars: boolean;
  includeResourceCalendars: boolean;
  delegatedAccess: string[];
  impersonationAccounts: string[];
}

/**
 * Teams integration configuration
 */
export interface ITeamsIntegration {
  enabled: boolean;
  includeChannelMeetings: boolean;
  includePrivateMeetings: boolean;
  includeRecordings: boolean;
  teamIds: string[];
  channelIds: string[];
}

/**
 * Outlook integration configuration
 */
export interface IOutlookIntegration {
  enabled: boolean;
  syncCategories: boolean;
  syncReminders: boolean;
  syncAttachments: boolean;
  addinEnabled: boolean;
}

/**
 * Power Platform integration configuration
 */
export interface IPowerPlatformIntegration {
  powerAutomate: {
    enabled: boolean;
    flowIds: string[];
  };
  powerBI: {
    enabled: boolean;
    reportIds: string[];
    datasetIds: string[];
  };
  powerApps: {
    enabled: boolean;
    appIds: string[];
  };
}

/**
 * Custom field mapping interface
 */
export interface ICustomFieldMapping {
  sourceField: string;
  targetField: string;
  dataType: 'string' | 'number' | 'date' | 'boolean' | 'choice';
  defaultValue?: any;
  isRequired: boolean;
  transformation?: string;
}

/**
 * User preferences interface
 */
export interface IUserPreferences {
  userId: string;
  defaultView: CalendarViewType;
  timeZone: string;
  workingHours: IWorkingHours;
  dateFormat: string;
  timeFormat: string;
  firstDayOfWeek: number;
  showWeekNumbers: boolean;
  selectedCalendars: string[];
  hiddenCalendars: string[];
  favoriteCalendars: string[];
  notifications: INotificationPreferences;
  privacy: IPrivacyPreferences;
  accessibility: IAccessibilityPreferences;
}

/**
 * Working hours interface
 */
export interface IWorkingHours {
  timeZone: string;
  monday: IWorkingDay;
  tuesday: IWorkingDay;
  wednesday: IWorkingDay;
  thursday: IWorkingDay;
  friday: IWorkingDay;
  saturday: IWorkingDay;
  sunday: IWorkingDay;
}

/**
 * Working day interface
 */
export interface IWorkingDay {
  isWorkingDay: boolean;
  startTime: string; // HH:mm format
  endTime: string;   // HH:mm format
  breaks: IBreakTime[];
}

/**
 * Break time interface
 */
export interface IBreakTime {
  startTime: string;
  endTime: string;
  name: string;
}

/**
 * Notification preferences interface
 */
export interface INotificationPreferences {
  enabled: boolean;
  emailNotifications: boolean;
  browserNotifications: boolean;
  mobileNotifications: boolean;
  reminderTimes: number[]; // minutes before event
  digestFrequency: 'none' | 'daily' | 'weekly';
  quietHours: {
    enabled: boolean;
    startTime: string;
    endTime: string;
  };
  types: {
    eventCreated: boolean;
    eventUpdated: boolean;
    eventDeleted: boolean;
    eventReminder: boolean;
    meetingInvitation: boolean;
    meetingResponse: boolean;
    calendarShared: boolean;
  };
}

/**
 * Privacy preferences interface
 */
export interface IPrivacyPreferences {
  sharePresence: boolean;
  shareFreeBusy: boolean;
  shareCalendarDetails: boolean;
  allowMeetingForwarding: boolean;
  hidePrivateEvents: boolean;
  anonymizeEventTitles: boolean;
  dataRetentionDays: number;
}

/**
 * Accessibility preferences interface
 */
export interface IAccessibilityPreferences {
  highContrast: boolean;
  largeText: boolean;
  reduceMotion: boolean;
  screenReaderOptimized: boolean;
  keyboardNavigation: boolean;
  colorBlindFriendly: boolean;
  focusIndicators: boolean;
}

/**
 * Performance configuration interface
 */
export interface IPerformanceConfiguration {
  caching: ICachingConfiguration;
  virtualization: IVirtualizationConfiguration;
  loading: ILoadingConfiguration;
  optimization: IOptimizationConfiguration;
}

/**
 * Caching configuration interface
 */
export interface ICachingConfiguration {
  enabled: boolean;
  defaultTtlMinutes: number;
  maxCacheSize: number;
  strategies: {
    events: 'memory' | 'session' | 'local' | 'none';
    calendars: 'memory' | 'session' | 'local' | 'none';
    metadata: 'memory' | 'session' | 'local' | 'none';
  };
}

/**
 * Virtualization configuration interface
 */
export interface IVirtualizationConfiguration {
  enabled: boolean;
  itemHeight: number;
  bufferSize: number;
  threshold: number;
}

/**
 * Loading configuration interface
 */
export interface ILoadingConfiguration {
  showProgress: boolean;
  showShimmer: boolean;
  lazyLoading: boolean;
  batchSize: number;
  timeout: number;
  retryAttempts: number;
  retryDelay: number;
}

/**
 * Optimization configuration interface
 */
export interface IOptimizationConfiguration {
  bundleOptimization: boolean;
  imageOptimization: boolean;
  compressionEnabled: boolean;
  minificationEnabled: boolean;
  treeShaking: boolean;
  codesplitting: boolean;
}

/**
 * Security configuration interface
 */
export interface ISecurityConfiguration {
  authentication: IAuthenticationConfiguration;
  authorization: IAuthorizationConfiguration;
  dataProtection: IDataProtectionConfiguration;
  auditLog: IAuditLogConfiguration;
}

/**
 * Authentication configuration interface
 */
export interface IAuthenticationConfiguration {
  provider: 'aad' | 'adfs' | 'forms' | 'anonymous';
  requireAuthentication: boolean;
  sessionTimeout: number;
  multiFactorRequired: boolean;
  allowedDomains: string[];
  blockedDomains: string[];
}

/**
 * Authorization configuration interface
 */
export interface IAuthorizationConfiguration {
  roleBasedAccess: boolean;
  roles: IRole[];
  permissions: IPermission[];
  defaultRole: string;
  inheritPermissions: boolean;
}

/**
 * Role interface
 */
export interface IRole {
  id: string;
  name: string;
  description: string;
  permissions: string[];
  isDefault: boolean;
  isSystem: boolean;
}

/**
 * Permission interface
 */
export interface IPermission {
  id: string;
  name: string;
  description: string;
  resource: string;
  action: string;
  scope: 'global' | 'site' | 'calendar' | 'event';
}

/**
 * Data protection configuration interface
 */
export interface IDataProtectionConfiguration {
  encryptionEnabled: boolean;
  encryptionAlgorithm: string;
  dataClassification: boolean;
  retentionPolicy: IRetentionPolicy;
  privacyCompliance: IPrivacyCompliance;
}

/**
 * Retention policy interface
 */
export interface IRetentionPolicy {
  enabled: boolean;
  defaultRetentionDays: number;
  eventRetentionDays: number;
  logRetentionDays: number;
  autoCleanup: boolean;
}

/**
 * Privacy compliance interface
 */
export interface IPrivacyCompliance {
  gdprCompliant: boolean;
  ccpaCompliant: boolean;
  dataSubjectRights: boolean;
  consentManagement: boolean;
  dataAnonymization: boolean;
}

/**
 * Audit log configuration interface
 */
export interface IAuditLogConfiguration {
  enabled: boolean;
  logLevel: 'error' | 'warn' | 'info' | 'debug';
  retentionDays: number;
  includeSensitiveData: boolean;
  events: {
    login: boolean;
    logout: boolean;
    calendarAccess: boolean;
    eventView: boolean;
    eventCreate: boolean;
    eventUpdate: boolean;
    eventDelete: boolean;
    export: boolean;
    share: boolean;
    configChange: boolean;
  };
}

/**
 * Analytics configuration interface
 */
export interface IAnalyticsConfiguration {
  enabled: boolean;
  provider: 'google' | 'adobe' | 'custom' | 'none';
  trackingId: string;
  events: {
    pageViews: boolean;
    calendarViews: boolean;
    eventClicks: boolean;
    searches: boolean;
    exports: boolean;
    errors: boolean;
  };
  customDimensions: ICustomDimension[];
  privacyMode: boolean;
}

/**
 * Custom dimension interface
 */
export interface ICustomDimension {
  id: string;
  name: string;
  scope: 'hit' | 'session' | 'user';
  value: string;
}

/**
 * Deployment configuration interface
 */
export interface IDeploymentConfiguration {
  environment: 'development' | 'staging' | 'production';
  version: string;
  buildNumber: string;
  deploymentDate: Date;
  features: IFeatureFlags;
  endpoints: IEndpointConfiguration;
  monitoring: IMonitoringConfiguration;
}

/**
 * Feature flags interface
 */
export interface IFeatureFlags {
  [featureName: string]: boolean;
}

/**
 * Endpoint configuration interface
 */
export interface IEndpointConfiguration {
  sharePointApi: string;
  graphApi: string;
  customApi?: string;
  timeouts: {
    default: number;
    sharePoint: number;
    graph: number;
  };
  retryPolicies: {
    maxRetries: number;
    backoffMultiplier: number;
    initialDelay: number;
  };
}

/**
 * Monitoring configuration interface
 */
export interface IMonitoringConfiguration {
  enabled: boolean;
  logLevel: 'error' | 'warn' | 'info' | 'debug' | 'trace';
  performanceTracking: boolean;
  errorTracking: boolean;
  userTracking: boolean;
  customMetrics: boolean;
  alerting: IAlertingConfiguration;
}

/**
 * Alerting configuration interface
 */
export interface IAlertingConfiguration {
  enabled: boolean;
  errorThreshold: number;
  performanceThreshold: number;
  recipients: string[];
  channels: ('email' | 'teams' | 'slack')[];
}

/**
 * Configuration validation result interface
 */
export interface IConfigurationValidationResult {
  isValid: boolean;
  errors: IConfigurationError[];
  warnings: IConfigurationWarning[];
}

/**
 * Configuration error interface
 */
export interface IConfigurationError {
  path: string;
  message: string;
  code: string;
  severity: 'error' | 'warning';
}

/**
 * Configuration warning interface
 */
export interface IConfigurationWarning {
  path: string;
  message: string;
  code: string;
  recommendation?: string;
}