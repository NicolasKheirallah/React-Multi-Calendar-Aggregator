import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';
import { IEventCreateRequest, IEventUpdateRequest } from '../models/IEventModels';
import { AppConstants } from '../constants/AppConstants';

export interface IValidationResult {
  isValid: boolean;
  errors: IValidationError[];
  warnings: IValidationWarning[];
}

export interface IValidationError {
  field: string;
  message: string;
  code: string;
  value?: unknown;
}

export interface IValidationWarning {
  field: string;
  message: string;
  code: string;
  value?: unknown;
}

export class ValidationUtils {
  /**
   * Validate email address format
   */
  public static isValidEmail(email: string): boolean {
    if (!email || typeof email !== 'string') return false;
    return AppConstants.REGEX.EMAIL.test(email.trim());
  }

  /**
   * Validate GUID format
   */
  public static isValidGuid(guid: string): boolean {
    if (!guid || typeof guid !== 'string') return false;
    return AppConstants.REGEX.GUID.test(guid.trim());
  }

  /**
   * Validate hex color format
   */
  public static isValidHexColor(color: string): boolean {
    if (!color || typeof color !== 'string') return false;
    return AppConstants.REGEX.HEX_COLOR.test(color.trim());
  }

  /**
   * Validate time format (24-hour)
   */
  public static isValidTime24H(time: string): boolean {
    if (!time || typeof time !== 'string') return false;
    return AppConstants.REGEX.TIME_24H.test(time.trim());
  }

  /**
   * Validate URL format
   */
  public static isValidUrl(url: string): boolean {
    if (!url || typeof url !== 'string') return false;
    try {
      void new URL(url.trim());
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Validate calendar event
   */
  public static validateCalendarEvent(event: ICalendarEvent): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!event) {
      errors.push({
        field: 'event',
        message: 'Event object is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Required field validation
    if (!event.title || event.title.trim().length === 0) {
      errors.push({
        field: 'title',
        message: 'Event title is required',
        code: 'REQUIRED_FIELD',
        value: event.title
      });
    }

    if (!event.start) {
      errors.push({
        field: 'start',
        message: 'Event start date is required',
        code: 'REQUIRED_FIELD',
        value: event.start
      });
    }

    if (!event.end) {
      errors.push({
        field: 'end',
        message: 'Event end date is required',
        code: 'REQUIRED_FIELD',
        value: event.end
      });
    }

    // Date validation
    if (event.start && event.end) {
      if (!(event.start instanceof Date) || isNaN(event.start.getTime())) {
        errors.push({
          field: 'start',
          message: 'Invalid start date format',
          code: 'INVALID_DATE_FORMAT',
          value: event.start
        });
      }

      if (!(event.end instanceof Date) || isNaN(event.end.getTime())) {
        errors.push({
          field: 'end',
          message: 'Invalid end date format',
          code: 'INVALID_DATE_FORMAT',
          value: event.end
        });
      }

      if (event.start instanceof Date && event.end instanceof Date && 
          !isNaN(event.start.getTime()) && !isNaN(event.end.getTime())) {
        
        if (event.start >= event.end) {
          errors.push({
            field: 'end',
            message: 'End date must be after start date',
            code: 'INVALID_DATE_RANGE',
            value: { start: event.start, end: event.end }
          });
        }

        // Check for extremely long events (more than 30 days)
        const durationMs = event.end.getTime() - event.start.getTime();
        const durationDays = durationMs / (1000 * 60 * 60 * 24);
        if (durationDays > 30) {
          warnings.push({
            field: 'duration',
            message: 'Event duration is longer than 30 days',
            code: 'LONG_DURATION',
            value: durationDays
          });
        }

        // Check for events in the past
        const now = new Date();
        if (event.end < now) {
          warnings.push({
            field: 'end',
            message: 'Event is in the past',
            code: 'PAST_EVENT',
            value: event.end
          });
        }

        // Check for very short events (less than 1 minute)
        if (durationMs < 60000 && !event.isAllDay) {
          warnings.push({
            field: 'duration',
            message: 'Event duration is less than 1 minute',
            code: 'SHORT_DURATION',
            value: durationMs / 60000
          });
        }
      }
    }

    // String length validation
    if (event.title && event.title.length > AppConstants.EVENT_TITLE_MAX_LENGTH) {
      errors.push({
        field: 'title',
        message: `Title exceeds maximum length of ${AppConstants.EVENT_TITLE_MAX_LENGTH} characters`,
        code: 'MAX_LENGTH_EXCEEDED',
        value: event.title.length
      });
    }

    if (event.description && event.description.length > AppConstants.EVENT_DESCRIPTION_MAX_LENGTH) {
      warnings.push({
        field: 'description',
        message: `Description exceeds recommended length of ${AppConstants.EVENT_DESCRIPTION_MAX_LENGTH} characters`,
        code: 'RECOMMENDED_LENGTH_EXCEEDED',
        value: event.description.length
      });
    }

    if (event.location && event.location.length > AppConstants.EVENT_LOCATION_MAX_LENGTH) {
      warnings.push({
        field: 'location',
        message: `Location exceeds recommended length of ${AppConstants.EVENT_LOCATION_MAX_LENGTH} characters`,
        code: 'RECOMMENDED_LENGTH_EXCEEDED',
        value: event.location.length
      });
    }

    // Calendar ID validation
    if (!event.calendarId || !this.isValidGuid(event.calendarId)) {
      errors.push({
        field: 'calendarId',
        message: 'Invalid calendar ID format',
        code: 'INVALID_FORMAT',
        value: event.calendarId
      });
    }

    // Attendee validation
    if (event.attendees && Array.isArray(event.attendees)) {
      event.attendees.forEach((attendee, index) => {
        if (!attendee.email || !this.isValidEmail(attendee.email)) {
          errors.push({
            field: `attendees[${index}].email`,
            message: 'Invalid attendee email address',
            code: 'INVALID_EMAIL',
            value: attendee.email
          });
        }

        if (!attendee.name || attendee.name.trim().length === 0) {
          warnings.push({
            field: `attendees[${index}].name`,
            message: 'Attendee name is empty',
            code: 'EMPTY_NAME',
            value: attendee.name
          });
        }

        const validResponses = ['accepted', 'declined', 'tentative', 'none'];
        if (attendee.response && !validResponses.includes(attendee.response)) {
          errors.push({
            field: `attendees[${index}].response`,
            message: 'Invalid attendee response status',
            code: 'INVALID_RESPONSE_STATUS',
            value: attendee.response
          });
        }

        const validTypes = ['required', 'optional', 'resource'];
        if (attendee.type && !validTypes.includes(attendee.type)) {
          errors.push({
            field: `attendees[${index}].type`,
            message: 'Invalid attendee type',
            code: 'INVALID_ATTENDEE_TYPE',
            value: attendee.type
          });
        }
      });

      // Check for too many attendees
      if (event.attendees.length > 100) {
        warnings.push({
          field: 'attendees',
          message: 'Event has more than 100 attendees',
          code: 'TOO_MANY_ATTENDEES',
          value: event.attendees.length
        });
      }

      // Check for duplicate email addresses
      const emails = event.attendees.map(a => a.email?.toLowerCase()).filter(Boolean);
      const uniqueEmails = new Set(emails);
      if (emails.length !== uniqueEmails.size) {
        warnings.push({
          field: 'attendees',
          message: 'Duplicate attendee email addresses found',
          code: 'DUPLICATE_ATTENDEES'
        });
      }
    }

    // URL validation
    if (event.webUrl && !this.isValidUrl(event.webUrl)) {
      warnings.push({
        field: 'webUrl',
        message: 'Invalid web URL format',
        code: 'INVALID_URL',
        value: event.webUrl
      });
    }

    // Importance validation
    if (event.importance) {
      const validImportance = ['low', 'normal', 'high'];
      if (!validImportance.includes(event.importance.toLowerCase())) {
        errors.push({
          field: 'importance',
          message: 'Invalid importance level',
          code: 'INVALID_IMPORTANCE',
          value: event.importance
        });
      }
    }

    // Sensitivity validation
    if (event.sensitivity) {
      const validSensitivity = ['normal', 'personal', 'private', 'confidential'];
      if (!validSensitivity.includes(event.sensitivity.toLowerCase())) {
        errors.push({
          field: 'sensitivity',
          message: 'Invalid sensitivity level',
          code: 'INVALID_SENSITIVITY',
          value: event.sensitivity
        });
      }
    }

    // Color validation
    if (event.color && !this.isValidHexColor(event.color)) {
      warnings.push({
        field: 'color',
        message: 'Invalid color format, should be hex color',
        code: 'INVALID_COLOR_FORMAT',
        value: event.color
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate event create request
   */
  public static validateEventCreateRequest(request: IEventCreateRequest): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!request) {
      errors.push({
        field: 'request',
        message: 'Create request object is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Required fields
    if (!request.title || request.title.trim().length === 0) {
      errors.push({
        field: 'title',
        message: 'Event title is required',
        code: 'REQUIRED_FIELD'
      });
    }

    if (!request.start) {
      errors.push({
        field: 'start',
        message: 'Start date is required',
        code: 'REQUIRED_FIELD'
      });
    }

    if (!request.end) {
      errors.push({
        field: 'end',
        message: 'End date is required',
        code: 'REQUIRED_FIELD'
      });
    }

    if (!request.calendarId) {
      errors.push({
        field: 'calendarId',
        message: 'Calendar ID is required',
        code: 'REQUIRED_FIELD'
      });
    }

    // Date validation
    if (request.start && request.end) {
      if (!(request.start instanceof Date) || isNaN(request.start.getTime())) {
        errors.push({
          field: 'start',
          message: 'Invalid start date format',
          code: 'INVALID_DATE_FORMAT'
        });
      }

      if (!(request.end instanceof Date) || isNaN(request.end.getTime())) {
        errors.push({
          field: 'end',
          message: 'Invalid end date format',
          code: 'INVALID_DATE_FORMAT'
        });
      }

      if (request.start instanceof Date && request.end instanceof Date &&
          !isNaN(request.start.getTime()) && !isNaN(request.end.getTime())) {
        
        if (request.start >= request.end) {
          errors.push({
            field: 'end',
            message: 'End date must be after start date',
            code: 'INVALID_DATE_RANGE'
          });
        }

        // Check if start date is too far in the future (more than 2 years)
        const now = new Date();
        const twoYearsFromNow = new Date(now.getFullYear() + 2, now.getMonth(), now.getDate());
        if (request.start > twoYearsFromNow) {
          warnings.push({
            field: 'start',
            message: 'Event is scheduled more than 2 years in the future',
            code: 'FAR_FUTURE_EVENT'
          });
        }

        // Check if event is in the past
        if (request.end < now) {
          warnings.push({
            field: 'end',
            message: 'Event is scheduled in the past',
            code: 'PAST_EVENT'
          });
        }
      }
    }

    // Calendar ID validation
    if (request.calendarId && !this.isValidGuid(request.calendarId)) {
      errors.push({
        field: 'calendarId',
        message: 'Invalid calendar ID format',
        code: 'INVALID_GUID'
      });
    }

    // Title validation
    if (request.title && request.title.length > AppConstants.EVENT_TITLE_MAX_LENGTH) {
      errors.push({
        field: 'title',
        message: `Title exceeds maximum length of ${AppConstants.EVENT_TITLE_MAX_LENGTH} characters`,
        code: 'MAX_LENGTH_EXCEEDED'
      });
    }

    // Attendee validation
    if (request.attendees && Array.isArray(request.attendees)) {
      request.attendees.forEach((attendee, index) => {
        if (!this.isValidEmail(attendee.email)) {
          errors.push({
            field: `attendees[${index}].email`,
            message: 'Invalid attendee email address',
            code: 'INVALID_EMAIL'
          });
        }
      });

      // Check for too many attendees
      if (request.attendees.length > 100) {
        warnings.push({
          field: 'attendees',
          message: 'Event has more than 100 attendees',
          code: 'TOO_MANY_ATTENDEES'
        });
      }
    }

    // Recurrence validation
    if (request.recurrence) {
      // Convert IRecurrencePattern to Record<string, unknown> for validation
      const recurrenceRecord: Record<string, unknown> = {
        type: request.recurrence.type,
        interval: request.recurrence.interval,
        daysOfWeek: request.recurrence.daysOfWeek,
        dayOfMonth: request.recurrence.dayOfMonth,
        weekOfMonth: request.recurrence.weekOfMonth,
        monthOfYear: request.recurrence.monthOfYear,
        endDate: request.recurrence.endDate,
        occurrences: request.recurrence.occurrences,
        firstDayOfWeek: request.recurrence.firstDayOfWeek
      };
      
      const recurrenceValidation = this.validateRecurrencePattern(recurrenceRecord);
      if (!recurrenceValidation.isValid) {
        recurrenceValidation.errors.forEach(error => {
          errors.push({
            field: `recurrence.${error.field}`,
            message: error.message,
            code: error.code
          });
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate event update request
   */
  public static validateEventUpdateRequest(request: IEventUpdateRequest): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!request) {
      errors.push({
        field: 'request',
        message: 'Update request object is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Required fields
    if (!request.id) {
      errors.push({
        field: 'id',
        message: 'Event ID is required for updates',
        code: 'REQUIRED_FIELD'
      });
    } else if (!this.isValidGuid(request.id)) {
      errors.push({
        field: 'id',
        message: 'Invalid event ID format',
        code: 'INVALID_GUID'
      });
    }

    // If dates are provided, validate them
    if (request.start && request.end) {
      if (!(request.start instanceof Date) || isNaN(request.start.getTime())) {
        errors.push({
          field: 'start',
          message: 'Invalid start date format',
          code: 'INVALID_DATE_FORMAT'
        });
      }

      if (!(request.end instanceof Date) || isNaN(request.end.getTime())) {
        errors.push({
          field: 'end',
          message: 'Invalid end date format',
          code: 'INVALID_DATE_FORMAT'
        });
      }

      if (request.start instanceof Date && request.end instanceof Date &&
          !isNaN(request.start.getTime()) && !isNaN(request.end.getTime())) {
        
        if (request.start >= request.end) {
          errors.push({
            field: 'end',
            message: 'End date must be after start date',
            code: 'INVALID_DATE_RANGE'
          });
        }
      }
    }

    // Validate title length if provided
    if (request.title !== undefined) {
      if (!request.title || request.title.trim().length === 0) {
        errors.push({
          field: 'title',
          message: 'Event title cannot be empty',
          code: 'INVALID_VALUE'
        });
      } else if (request.title.length > AppConstants.EVENT_TITLE_MAX_LENGTH) {
        errors.push({
          field: 'title',
          message: `Title exceeds maximum length of ${AppConstants.EVENT_TITLE_MAX_LENGTH} characters`,
          code: 'MAX_LENGTH_EXCEEDED'
        });
      }
    }

    // Validate calendar ID if provided
    if (request.calendarId !== undefined && !this.isValidGuid(request.calendarId)) {
      errors.push({
        field: 'calendarId',
        message: 'Invalid calendar ID format',
        code: 'INVALID_GUID'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate calendar source
   */
  public static validateCalendarSource(source: ICalendarSource): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!source) {
      errors.push({
        field: 'source',
        message: 'Calendar source object is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Required fields
    if (!source.id || !this.isValidGuid(source.id)) {
      errors.push({
        field: 'id',
        message: 'Valid calendar ID is required',
        code: 'INVALID_GUID'
      });
    }

    if (!source.title || source.title.trim().length === 0) {
      errors.push({
        field: 'title',
        message: 'Calendar title is required',
        code: 'REQUIRED_FIELD'
      });
    }

    if (!source.siteUrl || !this.isValidUrl(source.siteUrl)) {
      errors.push({
        field: 'siteUrl',
        message: 'Valid site URL is required',
        code: 'INVALID_URL'
      });
    }

    if (!source.type) {
      errors.push({
        field: 'type',
        message: 'Calendar type is required',
        code: 'REQUIRED_FIELD'
      });
    }

    // Color validation
    if (source.color && !this.isValidHexColor(source.color)) {
      warnings.push({
        field: 'color',
        message: 'Invalid color format, should be hex color',
        code: 'INVALID_COLOR_FORMAT'
      });
    }

    // URL validation
    if (source.url && !this.isValidUrl(source.url)) {
      warnings.push({
        field: 'url',
        message: 'Invalid calendar URL format',
        code: 'INVALID_URL'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate configuration settings
   */
  public static validateConfiguration(config: Record<string, unknown>): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!config) {
      errors.push({
        field: 'config',
        message: 'Configuration object is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Validate refresh interval
    if (config.refreshInterval !== undefined && config.refreshInterval !== null) {
      if (typeof config.refreshInterval !== 'number' || config.refreshInterval < AppConstants.VALIDATION.MIN_REFRESH_INTERVAL) {
        errors.push({
          field: 'refreshInterval',
          message: `Refresh interval must be at least ${AppConstants.VALIDATION.MIN_REFRESH_INTERVAL} minutes`,
          code: 'MIN_VALUE_VIOLATION'
        });
      }

      if (typeof config.refreshInterval === 'number' && config.refreshInterval > AppConstants.VALIDATION.MAX_REFRESH_INTERVAL) {
        warnings.push({
          field: 'refreshInterval',
          message: `Refresh interval exceeds recommended maximum of ${AppConstants.VALIDATION.MAX_REFRESH_INTERVAL} minutes`,
          code: 'MAX_VALUE_EXCEEDED'
        });
      }
    }

    // Validate max events
    if (config.maxEvents !== undefined && config.maxEvents !== null) {
      if (typeof config.maxEvents !== 'number' || config.maxEvents < AppConstants.VALIDATION.MIN_MAX_EVENTS) {
        errors.push({
          field: 'maxEvents',
          message: `Maximum events must be at least ${AppConstants.VALIDATION.MIN_MAX_EVENTS}`,
          code: 'MIN_VALUE_VIOLATION'
        });
      }

      if (typeof config.maxEvents === 'number' && config.maxEvents > AppConstants.VALIDATION.MAX_MAX_EVENTS) {
        warnings.push({
          field: 'maxEvents',
          message: `Maximum events exceeds recommended limit of ${AppConstants.VALIDATION.MAX_MAX_EVENTS}`,
          code: 'MAX_VALUE_EXCEEDED'
        });
      }
    }

    // Validate title
    if (config.title !== undefined) {
      if (typeof config.title !== 'string') {
        errors.push({
          field: 'title',
          message: 'Title must be a string',
          code: 'INVALID_TYPE'
        });
      } else {
        if (config.title.length < AppConstants.VALIDATION.MIN_TITLE_LENGTH) {
          errors.push({
            field: 'title',
            message: `Title must be at least ${AppConstants.VALIDATION.MIN_TITLE_LENGTH} character`,
            code: 'MIN_LENGTH_VIOLATION'
          });
        }

        if (config.title.length > AppConstants.VALIDATION.MAX_TITLE_LENGTH) {
          errors.push({
            field: 'title',
            message: `Title exceeds maximum length of ${AppConstants.VALIDATION.MAX_TITLE_LENGTH} characters`,
            code: 'MAX_LENGTH_EXCEEDED'
          });
        }
      }
    }

    // Validate view type
    if (config.viewType !== undefined) {
      const validViews = ['month', 'week', 'day', 'agenda', 'timeline'];
      if (!validViews.includes(config.viewType as string)) {
        errors.push({
          field: 'viewType',
          message: 'Invalid view type',
          code: 'INVALID_VIEW_TYPE'
        });
      }
    }

    // Validate selected calendars
    if (config.selectedCalendars !== undefined) {
      if (!Array.isArray(config.selectedCalendars)) {
        errors.push({
          field: 'selectedCalendars',
          message: 'Selected calendars must be an array',
          code: 'INVALID_TYPE'
        });
      } else {
        config.selectedCalendars.forEach((id: unknown, index: number) => {
          if (typeof id !== 'string' || !this.isValidGuid(id)) {
            errors.push({
              field: `selectedCalendars[${index}]`,
              message: 'Invalid calendar ID format',
              code: 'INVALID_GUID'
            });
          }
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate search query
   */
  public static validateSearchQuery(query: string): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!query || typeof query !== 'string') {
      errors.push({
        field: 'query',
        message: 'Search query must be a non-empty string',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    const trimmedQuery = query.trim();

    if (trimmedQuery.length === 0) {
      errors.push({
        field: 'query',
        message: 'Search query cannot be empty',
        code: 'REQUIRED_FIELD'
      });
    }

    if (trimmedQuery.length < 2) {
      warnings.push({
        field: 'query',
        message: 'Search queries with less than 2 characters may return too many results',
        code: 'SHORT_QUERY'
      });
    }

    if (trimmedQuery.length > 100) {
      warnings.push({
        field: 'query',
        message: 'Very long search queries may not perform well',
        code: 'LONG_QUERY'
      });
    }

    // Check for potentially dangerous characters
    const dangerousChars = ['<', '>', '"', "'", '&', 'script'];
    const lowerQuery = trimmedQuery.toLowerCase();
    const foundDangerous = dangerousChars.some(char => lowerQuery.includes(char));
    
    if (foundDangerous) {
      warnings.push({
        field: 'query',
        message: 'Search query contains potentially unsafe characters',
        code: 'UNSAFE_CHARACTERS'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate date range
   */
  public static validateDateRange(startDate: Date, endDate: Date): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!startDate) {
      errors.push({
        field: 'startDate',
        message: 'Start date is required',
        code: 'REQUIRED_FIELD'
      });
    } else if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
      errors.push({
        field: 'startDate',
        message: 'Invalid start date format',
        code: 'INVALID_DATE_FORMAT'
      });
    }

    if (!endDate) {
      errors.push({
        field: 'endDate',
        message: 'End date is required',
        code: 'REQUIRED_FIELD'
      });
    } else if (!(endDate instanceof Date) || isNaN(endDate.getTime())) {
      errors.push({
        field: 'endDate',
        message: 'Invalid end date format',
        code: 'INVALID_DATE_FORMAT'
      });
    }

    if (startDate instanceof Date && endDate instanceof Date && 
        !isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
      
      if (startDate >= endDate) {
        errors.push({
          field: 'endDate',
          message: 'End date must be after start date',
          code: 'INVALID_DATE_RANGE'
        });
      }

      // Check for very large date ranges (more than 1 year)
      const diffMs = endDate.getTime() - startDate.getTime();
      const diffDays = diffMs / (1000 * 60 * 60 * 24);
      
      if (diffDays > 365) {
        warnings.push({
          field: 'dateRange',
          message: 'Date range spans more than 1 year, which may affect performance',
          code: 'LARGE_DATE_RANGE'
        });
      }

      // Check for dates too far in the past or future
      const now = new Date();
      const twoYearsAgo = new Date(now.getFullYear() - 2, now.getMonth(), now.getDate());
      const twoYearsFromNow = new Date(now.getFullYear() + 2, now.getMonth(), now.getDate());

      if (endDate < twoYearsAgo) {
        warnings.push({
          field: 'endDate',
          message: 'Date range is more than 2 years in the past',
          code: 'OLD_DATE_RANGE'
        });
      }

      if (startDate > twoYearsFromNow) {
        warnings.push({
          field: 'startDate',
          message: 'Date range is more than 2 years in the future',
          code: 'FUTURE_DATE_RANGE'
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate time range
   */
  public static validateTimeRange(startTime: string, endTime: string): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!startTime) {
      errors.push({
        field: 'startTime',
        message: 'Start time is required',
        code: 'REQUIRED_FIELD'
      });
    } else if (!this.isValidTime24H(startTime)) {
      errors.push({
        field: 'startTime',
        message: 'Invalid start time format (use HH:mm)',
        code: 'INVALID_FORMAT'
      });
    }

    if (!endTime) {
      errors.push({
        field: 'endTime',
        message: 'End time is required',
        code: 'REQUIRED_FIELD'
      });
    } else if (!this.isValidTime24H(endTime)) {
      errors.push({
        field: 'endTime',
        message: 'Invalid end time format (use HH:mm)',
        code: 'INVALID_FORMAT'
      });
    }

    if (startTime && endTime && this.isValidTime24H(startTime) && this.isValidTime24H(endTime)) {
      const [startHour, startMinute] = startTime.split(':').map(Number);
      const [endHour, endMinute] = endTime.split(':').map(Number);
      
      const startMinutes = startHour * 60 + startMinute;
      const endMinutes = endHour * 60 + endMinute;

      if (startMinutes >= endMinutes) {
        errors.push({
          field: 'endTime',
          message: 'End time must be after start time',
          code: 'INVALID_TIME_RANGE'
        });
      }

      // Check for very short time ranges (less than 15 minutes)
      const durationMinutes = endMinutes - startMinutes;
      if (durationMinutes < 15) {
        warnings.push({
          field: 'timeRange',
          message: 'Time range is less than 15 minutes',
          code: 'SHORT_TIME_RANGE'
        });
      }

      // Check for very long time ranges (more than 12 hours)
      if (durationMinutes > 720) {
        warnings.push({
          field: 'timeRange',
          message: 'Time range is more than 12 hours',
          code: 'LONG_TIME_RANGE'
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate array of calendar IDs
   */
  public static validateCalendarIds(calendarIds: string[]): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!calendarIds) {
      errors.push({
        field: 'calendarIds',
        message: 'Calendar IDs array is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    if (!Array.isArray(calendarIds)) {
      errors.push({
        field: 'calendarIds',
        message: 'Calendar IDs must be an array',
        code: 'INVALID_TYPE'
      });
      return { isValid: false, errors, warnings };
    }

    if (calendarIds.length === 0) {
      warnings.push({
        field: 'calendarIds',
        message: 'No calendars selected',
        code: 'NO_CALENDARS_SELECTED'
      });
    }

    calendarIds.forEach((id, index) => {
      if (!id || typeof id !== 'string') {
        errors.push({
          field: `calendarIds[${index}]`,
          message: 'Calendar ID must be a non-empty string',
          code: 'INVALID_TYPE',
          value: id
        });
      } else if (!this.isValidGuid(id)) {
        errors.push({
          field: `calendarIds[${index}]`,
          message: 'Invalid calendar ID format',
          code: 'INVALID_GUID',
          value: id
        });
      }
    });

    // Check for duplicates
    const uniqueIds = new Set(calendarIds.filter(id => typeof id === 'string'));
    if (uniqueIds.size !== calendarIds.length) {
      warnings.push({
        field: 'calendarIds',
        message: 'Duplicate calendar IDs found',
        code: 'DUPLICATE_VALUES'
      });
    }

    // Check for too many calendars
    if (calendarIds.length > 20) {
      warnings.push({
        field: 'calendarIds',
        message: 'More than 20 calendars selected may affect performance',
        code: 'TOO_MANY_CALENDARS'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate pagination parameters
   */
  public static validatePagination(pageSize: number, currentPage: number): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (typeof pageSize !== 'number' || pageSize <= 0) {
      errors.push({
        field: 'pageSize',
        message: 'Page size must be a positive number',
        code: 'INVALID_PAGE_SIZE'
      });
    }

    if (pageSize > 1000) {
      warnings.push({
        field: 'pageSize',
        message: 'Large page sizes may affect performance',
        code: 'LARGE_PAGE_SIZE'
      });
    }

    if (typeof currentPage !== 'number' || currentPage < 0) {
      errors.push({
        field: 'currentPage',
        message: 'Current page must be 0 or greater',
        code: 'INVALID_CURRENT_PAGE'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate recurrence pattern
   */
  public static validateRecurrencePattern(pattern: Record<string, unknown>): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!pattern) {
      errors.push({
        field: 'pattern',
        message: 'Recurrence pattern is required',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Validate type
    const validTypes = ['daily', 'weekly', 'monthly', 'yearly', 'weekdays', 'custom'];
    if (!pattern.type || !validTypes.includes(pattern.type as string)) {
      errors.push({
        field: 'type',
        message: 'Invalid recurrence type',
        code: 'INVALID_RECURRENCE_TYPE'
      });
    }

    // Validate interval
    if (!pattern.interval || typeof pattern.interval !== 'number' || pattern.interval < 1) {
      errors.push({
        field: 'interval',
        message: 'Recurrence interval must be at least 1',
        code: 'INVALID_INTERVAL'
      });
    }

    if (pattern.interval && typeof pattern.interval === 'number' && pattern.interval > 999) {
      errors.push({
        field: 'interval',
        message: 'Recurrence interval cannot exceed 999',
        code: 'INTERVAL_TOO_LARGE'
      });
    }

    // Validate type-specific properties
    switch (pattern.type) {
      case 'weekly': {
        if (pattern.daysOfWeek) {
          if (!Array.isArray(pattern.daysOfWeek) || pattern.daysOfWeek.length === 0) {
            errors.push({
              field: 'daysOfWeek',
              message: 'Weekly recurrence must specify at least one day of the week',
              code: 'MISSING_DAYS_OF_WEEK'
            });
          } else {
            const invalidDays = pattern.daysOfWeek.filter((day: unknown) => 
              typeof day !== 'number' || day < 0 || day > 6
            );
            if (invalidDays.length > 0) {
              errors.push({
                field: 'daysOfWeek',
                message: 'Days of week must be numbers between 0 (Sunday) and 6 (Saturday)',
                code: 'INVALID_DAYS_OF_WEEK'
              });
            }
          }
        }
        break;
      }
      case 'monthly': {
        if (pattern.dayOfMonth && (typeof pattern.dayOfMonth !== 'number' || 
            pattern.dayOfMonth < 1 || pattern.dayOfMonth > 31)) {
          errors.push({
            field: 'dayOfMonth',
            message: 'Day of month must be between 1 and 31',
            code: 'INVALID_DAY_OF_MONTH'
          });
        }
        if (pattern.weekOfMonth && (typeof pattern.weekOfMonth !== 'number' || 
            (pattern.weekOfMonth < -1 || pattern.weekOfMonth === 0 || pattern.weekOfMonth > 4))) {
          errors.push({
            field: 'weekOfMonth',
            message: 'Week of month must be between 1-4 or -1 (last week)',
            code: 'INVALID_WEEK_OF_MONTH'
          });
        }
        break;
      }
      case 'yearly': {
        if (pattern.monthOfYear && (typeof pattern.monthOfYear !== 'number' || 
            pattern.monthOfYear < 1 || pattern.monthOfYear > 12)) {
          errors.push({
            field: 'monthOfYear',
            message: 'Month of year must be between 1 and 12',
            code: 'INVALID_MONTH_OF_YEAR'
          });
        }
        if (pattern.dayOfMonth && (typeof pattern.dayOfMonth !== 'number' || 
            pattern.dayOfMonth < 1 || pattern.dayOfMonth > 31)) {
          errors.push({
            field: 'dayOfMonth',
            message: 'Day of month must be between 1 and 31',
            code: 'INVALID_DAY_OF_MONTH'
          });
        }
        break;
      }
    }

    // Validate end conditions
    if (pattern.endDate && pattern.occurrences) {
      errors.push({
        field: 'endDate',
        message: 'Cannot specify both end date and number of occurrences',
        code: 'CONFLICTING_END_CONDITIONS'
      });
    }

    if (pattern.endDate && (!(pattern.endDate instanceof Date) || isNaN((pattern.endDate as Date).getTime()))) {
      errors.push({
        field: 'endDate',
        message: 'Invalid end date format',
        code: 'INVALID_END_DATE'
      });
    }

    if (pattern.occurrences && (typeof pattern.occurrences !== 'number' || pattern.occurrences < 1)) {
      errors.push({
        field: 'occurrences',
        message: 'Number of occurrences must be at least 1',
        code: 'INVALID_OCCURRENCES'
      });
    }

    if (pattern.occurrences && typeof pattern.occurrences === 'number' && pattern.occurrences > 999) {
      errors.push({
        field: 'occurrences',
        message: 'Number of occurrences cannot exceed 999',
        code: 'OCCURRENCES_TOO_LARGE'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Sanitize HTML content
   */
  public static sanitizeHtml(html: string): string {
    if (!html || typeof html !== 'string') return '';

    // Basic HTML sanitization - remove script tags and dangerous attributes
    return html
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
      .replace(/on\w+\s*=/gi, '')
      .replace(/<iframe/gi, '&lt;iframe')
      .replace(/<object/gi, '&lt;object')
      .replace(/<embed/gi, '&lt;embed')
      .replace(/<form/gi, '&lt;form')
      .replace(/<input/gi, '&lt;input')
      .replace(/<textarea/gi, '&lt;textarea')
      .replace(/<select/gi, '&lt;select')
      .replace(/<button/gi, '&lt;button');
  }

  /**
   * Validate and sanitize input string
   */
  public static sanitizeInput(input: string, maxLength?: number): string {
    if (!input || typeof input !== 'string') return '';

    let sanitized = input.trim();
    
    // Remove potentially dangerous characters
    sanitized = sanitized.replace(/[<>'"&]/g, (match) => {
      const entityMap: { [key: string]: string } = {
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#x27;',
        '&': '&amp;'
      };
      return entityMap[match];
    });

    // Remove control characters except tab, newline, and carriage return
    sanitized = sanitized.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');

    // Truncate if necessary
    if (maxLength && sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength);
    }

    return sanitized;
  }

  /**
   * Validate file upload
   */
  public static validateFileUpload(file: File, allowedTypes: string[], maxSizeBytes: number): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!file) {
      errors.push({
        field: 'file',
        message: 'No file selected',
        code: 'REQUIRED_FIELD'
      });
      return { isValid: false, errors, warnings };
    }

    // Check file type
    if (allowedTypes.length > 0 && !allowedTypes.includes(file.type)) {
      errors.push({
        field: 'file',
        message: `File type ${file.type} is not allowed. Allowed types: ${allowedTypes.join(', ')}`,
        code: 'INVALID_FILE_TYPE'
      });
    }

    // Check file size
    if (file.size > maxSizeBytes) {
      errors.push({
        field: 'file',
        message: `File size ${file.size} bytes exceeds maximum of ${maxSizeBytes} bytes`,
        code: 'FILE_TOO_LARGE'
      });
    }

    // Warn about large files
    if (file.size > maxSizeBytes * 0.8) {
      warnings.push({
        field: 'file',
        message: 'File is close to maximum size limit',
        code: 'LARGE_FILE'
      });
    }

    // Check file name
    if (file.name.length > 255) {
      errors.push({
        field: 'fileName',
        message: 'File name is too long (maximum 255 characters)',
        code: 'FILENAME_TOO_LONG'
      });
    }

    // Check for suspicious file names
    const suspiciousPatterns = [/\.exe$/i, /\.bat$/i, /\.cmd$/i, /\.scr$/i, /\.vbs$/i];
    if (suspiciousPatterns.some(pattern => pattern.test(file.name))) {
      warnings.push({
        field: 'fileName',
        message: 'File type may be potentially unsafe',
        code: 'SUSPICIOUS_FILE_TYPE'
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate permissions
   */
  public static validatePermissions(requiredPermissions: string[], userPermissions: string[]): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationWarning[] = [];

    if (!requiredPermissions || !Array.isArray(requiredPermissions)) {
      errors.push({
        field: 'requiredPermissions',
        message: 'Required permissions must be an array',
        code: 'INVALID_TYPE'
      });
      return { isValid: false, errors, warnings };
    }

    if (!userPermissions || !Array.isArray(userPermissions)) {
      errors.push({
        field: 'userPermissions',
        message: 'User permissions must be an array',
        code: 'INVALID_TYPE'
      });
      return { isValid: false, errors, warnings };
    }

    const missingPermissions = requiredPermissions.filter(permission => 
      !userPermissions.includes(permission)
    );

    if (missingPermissions.length > 0) {
      errors.push({
        field: 'permissions',
        message: `Missing required permissions: ${missingPermissions.join(', ')}`,
        code: 'INSUFFICIENT_PERMISSIONS',
        value: missingPermissions
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Batch validate multiple items
   */
  public static batchValidate<T>(
    items: T[],
    validator: (item: T, index: number) => IValidationResult
  ): { 
    isValid: boolean; 
    results: IValidationResult[]; 
    summary: { 
      totalItems: number; 
      validItems: number; 
      invalidItems: number; 
      totalErrors: number; 
      totalWarnings: number; 
    } 
  } {
    if (!items || !Array.isArray(items)) {
      return {
        isValid: false,
        results: [],
        summary: {
          totalItems: 0,
          validItems: 0,
          invalidItems: 0,
          totalErrors: 1,
          totalWarnings: 0
        }
      };
    }

    const results: IValidationResult[] = [];
    let totalErrors = 0;
    let totalWarnings = 0;
    let validItems = 0;

    items.forEach((item, index) => {
      try {
        const result = validator(item, index);
        results.push(result);
        
        if (result.isValid) {
          validItems++;
        }
        
        totalErrors += result.errors.length;
        totalWarnings += result.warnings.length;
      } catch (error) {
        const errorResult: IValidationResult = {
          isValid: false,
          errors: [{
            field: `item[${index}]`,
            message: `Validation failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
            code: 'VALIDATION_ERROR'
          }],
          warnings: []
        };
        results.push(errorResult);
        totalErrors++;
      }
    });

    return {
      isValid: totalErrors === 0,
      results,
      summary: {
        totalItems: items.length,
        validItems,
        invalidItems: items.length - validItems,
        totalErrors,
        totalWarnings
      }
    };
  }

  /**
   * Create validation error
   */
  public static createError(field: string, message: string, code: string, value?: unknown): IValidationError {
    return { field, message, code, value };
  }

  /**
   * Create validation warning
   */
  public static createWarning(field: string, message: string, code: string, value?: unknown): IValidationWarning {
    return { field, message, code, value };
  }

  /**
   * Merge validation results
   */
  public static mergeValidationResults(...results: IValidationResult[]): IValidationResult {
    const mergedErrors: IValidationError[] = [];
    const mergedWarnings: IValidationWarning[] = [];

    results.forEach(result => {
      if (result) {
        mergedErrors.push(...(result.errors || []));
        mergedWarnings.push(...(result.warnings || []));
      }
    });

    return {
      isValid: mergedErrors.length === 0,
      errors: mergedErrors,
      warnings: mergedWarnings
    };
  }

  /**
   * Format validation result for display
   */
  public static formatValidationResult(result: IValidationResult): string {
    if (!result) return 'No validation result provided';

    const messages: string[] = [];
    
    if (result.errors.length > 0) {
      messages.push('Errors:');
      result.errors.forEach(error => {
        messages.push(`  - ${error.field}: ${error.message} (${error.code})`);
      });
    }

    if (result.warnings.length > 0) {
      messages.push('Warnings:');
      result.warnings.forEach(warning => {
        messages.push(`  - ${warning.field}: ${warning.message} (${warning.code})`);
      });
    }

    if (messages.length === 0) {
      messages.push('Validation passed successfully');
    }

    return messages.join('\n');
  }

  /**
   * Check if validation result has specific error code
   */
  public static hasErrorCode(result: IValidationResult, code: string): boolean {
    return result.errors.some(error => error.code === code);
  }

  /**
   * Check if validation result has specific warning code
   */
  public static hasWarningCode(result: IValidationResult, code: string): boolean {
    return result.warnings.some(warning => warning.code === code);
  }

  /**
   * Get errors for specific field
   */
  public static getFieldErrors(result: IValidationResult, field: string): IValidationError[] {
    return result.errors.filter(error => error.field === field);
  }

  /**
   * Get warnings for specific field
   */
  public static getFieldWarnings(result: IValidationResult, field: string): IValidationWarning[] {
    return result.warnings.filter(warning => warning.field === field);
  }

  /**
   * Convert validation result to summary object
   */
  public static toSummary(result: IValidationResult): {
    isValid: boolean;
    errorCount: number;
    warningCount: number;
    errorCodes: string[];
    warningCodes: string[];
  } {
    return {
      isValid: result.isValid,
      errorCount: result.errors.length,
      warningCount: result.warnings.length,
      errorCodes: result.errors.map(e => e.code),
      warningCodes: result.warnings.map(w => w.code)
    };
  }
}