/**
 * Base filter interface
 */
export interface IBaseFilter {
  id: string;
  name: string;
  isEnabled: boolean;
  isDefault?: boolean;
}

/**
 * Date range filter interface
 */
export interface IDateRangeFilter extends IBaseFilter {
  type: 'dateRange';
  startDate: Date;
  endDate: Date;
  preset?: 'today' | 'tomorrow' | 'thisWeek' | 'nextWeek' | 'thisMonth' | 'nextMonth' | 'custom';
  includeAllDay?: boolean;
  includeRecurring?: boolean;
}

/**
 * Calendar filter interface
 */
export interface ICalendarFilter extends IBaseFilter {
  type: 'calendar';
  calendarIds: string[];
  includeAllCalendars: boolean;
  excludeCalendarIds?: string[];
}

/**
 * Category filter interface
 */
export interface ICategoryFilter extends IBaseFilter {
  type: 'category';
  categories: string[];
  includeUncategorized: boolean;
  matchType: 'exact' | 'contains' | 'startsWith';
}

/**
 * Text search filter interface
 */
export interface ITextSearchFilter extends IBaseFilter {
  type: 'textSearch';
  query: string;
  searchFields: ('title' | 'description' | 'location' | 'organizer' | 'attendees')[];
  matchType: 'contains' | 'exact' | 'startsWith' | 'regex';
  caseSensitive: boolean;
}

/**
 * Attendee filter interface
 */
export interface IAttendeeFilter extends IBaseFilter {
  type: 'attendee';
  attendees: string[];
  includeOrganizer: boolean;
  responseStatus?: ('accepted' | 'declined' | 'tentative' | 'none')[];
  attendeeType?: ('required' | 'optional' | 'resource')[];
}

/**
 * Time range filter interface
 */
export interface ITimeRangeFilter extends IBaseFilter {
  type: 'timeRange';
  startTime: string; // HH:mm format
  endTime: string;   // HH:mm format
  daysOfWeek: number[]; // 0 = Sunday, 1 = Monday, etc.
  includeAllDay: boolean;
}

/**
 * Location filter interface
 */
export interface ILocationFilter extends IBaseFilter {
  type: 'location';
  locations: string[];
  includeOnlineEvents: boolean;
  includeEventsWithoutLocation: boolean;
  matchType: 'exact' | 'contains' | 'startsWith';
}

/**
 * Importance filter interface
 */
export interface IImportanceFilter extends IBaseFilter {
  type: 'importance';
  importanceLevels: ('high' | 'normal' | 'low')[];
}

/**
 * Sensitivity filter interface
 */
export interface ISensitivityFilter extends IBaseFilter {
  type: 'sensitivity';
  sensitivityLevels: ('normal' | 'personal' | 'private' | 'confidential')[];
}

/**
 * Show as filter interface
 */
export interface IShowAsFilter extends IBaseFilter {
  type: 'showAs';
  showAsTypes: ('free' | 'tentative' | 'busy' | 'outOfOffice' | 'workingElsewhere')[];
}

/**
 * Duration filter interface
 */
export interface IDurationFilter extends IBaseFilter {
  type: 'duration';
  minDuration?: number; // minutes
  maxDuration?: number; // minutes
  includeAllDay: boolean;
}

/**
 * Recurrence filter interface
 */
export interface IRecurrenceFilter extends IBaseFilter {
  type: 'recurrence';
  includeRecurring: boolean;
  includeNonRecurring: boolean;
  recurrenceTypes?: ('daily' | 'weekly' | 'monthly' | 'yearly' | 'custom')[];
}

/**
 * Custom property filter interface
 */
export interface ICustomPropertyFilter extends IBaseFilter {
  type: 'customProperty';
  propertyName: string;
  propertyValue: any;
  operator: 'equals' | 'contains' | 'startsWith' | 'endsWith' | 'greaterThan' | 'lessThan' | 'exists';
}

/**
 * Union type for all filter types
 */
export type IEventFilter = 
  | IDateRangeFilter
  | ICalendarFilter
  | ICategoryFilter
  | ITextSearchFilter
  | IAttendeeFilter
  | ITimeRangeFilter
  | ILocationFilter
  | IImportanceFilter
  | ISensitivityFilter
  | IShowAsFilter
  | IDurationFilter
  | IRecurrenceFilter
  | ICustomPropertyFilter;

/**
 * Filter group interface
 */
export interface IFilterGroup {
  id: string;
  name: string;
  filters: IEventFilter[];
  operator: 'AND' | 'OR';
  isEnabled: boolean;
  description?: string;
}

/**
 * Filter set interface
 */
export interface IFilterSet {
  id: string;
  name: string;
  description?: string;
  filterGroups: IFilterGroup[];
  globalOperator: 'AND' | 'OR';
  isDefault: boolean;
  isSystem: boolean;
  createdBy: string;
  createdDate: Date;
  modifiedDate: Date;
  tags?: string[];
}

/**
 * Filter preset interface
 */
export interface IFilterPreset {
  id: string;
  name: string;
  description: string;
  icon?: string;
  filterSet: IFilterSet;
  isQuickFilter: boolean;
  sortOrder: number;
}

/**
 * Quick filter interface
 */
export interface IQuickFilter {
  id: string;
  label: string;
  icon: string;
  tooltip: string;
  filter: IEventFilter;
  isActive: boolean;
  count?: number;
}

/**
 * Advanced search criteria interface
 */
export interface IAdvancedSearchCriteria {
  filterSets: IFilterSet[];
  sortBy: ISortOption[];
  groupBy?: IGroupByOption;
  pagination?: IPaginationOption;
}

/**
 * Sort option interface
 */
export interface ISortOption {
  field: 'title' | 'start' | 'end' | 'created' | 'modified' | 'calendar' | 'category' | 'importance';
  direction: 'asc' | 'desc';
  priority: number;
}

/**
 * Group by option interface
 */
export interface IGroupByOption {
  field: 'calendar' | 'category' | 'date' | 'organizer' | 'importance' | 'location';
  sortDirection: 'asc' | 'desc';
  showEmptyGroups: boolean;
}

/**
 * Pagination option interface
 */
export interface IPaginationOption {
  pageSize: number;
  currentPage: number;
  totalItems?: number;
}

/**
 * Filter validation result interface
 */
export interface IFilterValidationResult {
  isValid: boolean;
  errors: IFilterValidationError[];
  warnings: IFilterValidationWarning[];
}

/**
 * Filter validation error interface
 */
export interface IFilterValidationError {
  filterId: string;
  field: string;
  message: string;
  code: string;
}

/**
 * Filter validation warning interface
 */
export interface IFilterValidationWarning {
  filterId: string;
  field: string;
  message: string;
  code: string;
}

/**
 * Filter statistics interface
 */
export interface IFilterStatistics {
  totalFilters: number;
  activeFilters: number;
  filtersByType: { [type: string]: number };
  mostUsedFilters: { filterId: string; usageCount: number }[];
  averageFilterComplexity: number;
}

/**
 * Saved search interface
 */
export interface ISavedSearch {
  id: string;
  name: string;
  description?: string;
  criteria: IAdvancedSearchCriteria;
  isPublic: boolean;
  createdBy: string;
  createdDate: Date;
  lastUsed?: Date;
  usageCount: number;
  tags?: string[];
}

/**
 * Search suggestion interface
 */
export interface ISearchSuggestion {
  text: string;
  type: 'recent' | 'popular' | 'autocomplete';
  score: number;
  metadata?: any;
}

/**
 * Filter application result interface
 */
export interface IFilterApplicationResult<T> {
  items: T[];
  totalCount: number;
  appliedFilters: IEventFilter[];
  statistics: {
    totalItems: number;
    filteredItems: number;
    filterDuration: number;
  };
}

/**
 * Dynamic filter interface
 */
export interface IDynamicFilter extends IBaseFilter {
  type: 'dynamic';
  query: string; // Dynamic query that gets evaluated at runtime
  parameters?: { [key: string]: any };
  refreshInterval?: number; // minutes
}

/**
 * Conditional filter interface
 */
export interface IConditionalFilter extends IBaseFilter {
  type: 'conditional';
  condition: string; // Boolean expression
  trueFilter: IEventFilter;
  falseFilter?: IEventFilter;
}

/**
 * Smart filter interface
 */
export interface ISmartFilter extends IBaseFilter {
  type: 'smart';
  algorithm: 'ml' | 'rules' | 'heuristic';
  configuration: any;
  learningData?: any;
  confidence?: number;
}

/**
 * Filter template interface
 */
export interface IFilterTemplate {
  id: string;
  name: string;
  description: string;
  category: string;
  filterSet: Partial<IFilterSet>;
  variables?: IFilterVariable[];
}

/**
 * Filter variable interface
 */
export interface IFilterVariable {
  name: string;
  type: 'string' | 'number' | 'date' | 'boolean' | 'list';
  defaultValue?: any;
  description?: string;
  validation?: IFilterVariableValidation;
}

/**
 * Filter variable validation interface
 */
export interface IFilterVariableValidation {
  required?: boolean;
  minLength?: number;
  maxLength?: number;
  pattern?: string;
  options?: any[];
}

/**
 * Export filters for external use
 */
export * from './ICalendarModels';