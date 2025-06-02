import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICalendarEvent, ICalendarSource, CalendarSourceType, ICalendarService, IEventAttachment, IEventAttendee } from '../models/ICalendarModels';
import { IEventCreateRequest, IEventUpdateRequest, IEventSearchCriteria, IEventSearchResult, IExtendedCalendarEvent } from '../models/IEventModels';
import { SharePointCalendarService } from './SharePointCalendarService';
import { ExchangeCalendarService } from './ExchangeCalendarService';
import { CacheService } from './CacheService';
import { AppConstants } from '../constants/AppConstants';
import { ValidationUtils } from '../utils/ValidationUtils';
import { DateUtils } from '../utils/DateUtils';
import { ColorUtils } from '../utils/ColorUtils';

export class CalendarService implements ICalendarService {
  private sharePointService: SharePointCalendarService;
  private exchangeService: ExchangeCalendarService;
  private cacheService: CacheService;
  private isInitialized: boolean = false;

  constructor(context: WebPartContext) {
    this.sharePointService = new SharePointCalendarService(context);
    this.exchangeService = new ExchangeCalendarService(context);
    this.cacheService = CacheService.getInstance();
  }

  /**
   * Initialize the calendar service
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Configure cache settings
      this.cacheService.configure({
        defaultTtlMinutes: AppConstants.CACHE_DURATION_MINUTES
      });

      this.isInitialized = true;
    } catch (error) {
      console.error('Failed to initialize CalendarService:', error);
      throw new Error(AppConstants.ERROR_MESSAGES.GENERAL_ERROR);
    }
  }

  /**
   * Get all available calendar sources
   */
  public async getCalendarSources(includeExchange: boolean = false): Promise<ICalendarSource[]> {
    try {
      await this.initialize();

      const cachedSources = this.cacheService.getCachedSources();
      if (cachedSources) {
        return includeExchange 
          ? cachedSources 
          : cachedSources.filter(s => s.type === CalendarSourceType.SharePoint);
      }

      const allSources: ICalendarSource[] = [];

      // Get SharePoint calendar sources
      try {
        const sharePointSources = await this.sharePointService.getSharePointCalendars();
        allSources.push(...sharePointSources);
      } catch (error) {
        console.warn('Failed to get SharePoint calendars:', error);
        // Continue without SharePoint calendars
      }

      // Get Exchange calendar sources if enabled
      if (includeExchange) {
        try {
          const exchangeSources = await this.exchangeService.getExchangeCalendars();
          allSources.push(...exchangeSources);
        } catch (error) {
          console.warn('Failed to get Exchange calendars:', error);
          // Continue without Exchange calendars
        }
      }

      // Validate and enhance sources
      const validatedSources = await this.validateAndEnhanceSources(allSources);

      // Cache the results
      this.cacheService.setCachedSources(validatedSources, AppConstants.CACHE_DURATION_MINUTES);

      return validatedSources;
    } catch (error) {
      console.error('Error getting calendar sources:', error);
      throw new Error(AppConstants.ERROR_MESSAGES.SHAREPOINT_API_ERROR);
    }
  }

  /**
   * Get events from a specific calendar source
   */
  public async getEventsFromSource(source: ICalendarSource, maxEvents: number = 100): Promise<ICalendarEvent[]> {
    try {
      // Validate source
      const validation = ValidationUtils.validateCalendarSource(source);
      if (!validation.isValid) {
        throw new Error(`Invalid calendar source: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Check cache first
      const cachedEvents = this.cacheService.getCachedEvents(source.id);
      if (cachedEvents && cachedEvents.length <= maxEvents) {
        return cachedEvents.slice(0, maxEvents);
      }

      let events: ICalendarEvent[] = [];

      // Route to appropriate service based on source type
      switch (source.type) {
        case CalendarSourceType.SharePoint: {
          events = await this.sharePointService.getEventsFromCalendar(source, maxEvents);
          break;
        }
        case CalendarSourceType.Exchange: {
          events = await this.exchangeService.getEventsFromCalendar(source, maxEvents);
          break;
        }
        default: {
          throw new Error(`Unsupported calendar source type: ${source.type}`);
        }
      }

      // Validate and enhance events
      events = await this.validateAndEnhanceEvents(events, source);

      // Cache the results
      this.cacheService.setCachedEvents(source.id, events, AppConstants.CACHE_DURATION_MINUTES);

      return events.slice(0, maxEvents);
    } catch (error) {
      console.error(`Error getting events from ${source.title}:`, error);
      throw error;
    }
  }

  /**
   * Get events from multiple calendar sources
   */
  public async getEventsFromSources(sources: ICalendarSource[], maxEvents: number = 1000): Promise<ICalendarEvent[]> {
    try {
      // Check for cached aggregated events
      const sourceIds = sources.map(s => s.id);
      const cachedEvents = this.cacheService.getCachedAggregatedEvents(sourceIds);
      if (cachedEvents && cachedEvents.length <= maxEvents) {
        return cachedEvents.slice(0, maxEvents);
      }

      const allEvents: ICalendarEvent[] = [];
      const errors: string[] = [];

      // Process sources in parallel with error handling
      const eventPromises = sources.map(async (source) => {
        try {
          const events = await this.getEventsFromSource(source, Math.ceil(maxEvents / sources.length));
          return events;
        } catch (error) {
          console.error(`Failed to get events from ${source.title}:`, error);
          errors.push(`${source.title}: ${error instanceof Error ? error.message : 'Unknown error'}`);
          return [];
        }
      });

      const eventArrays = await Promise.all(eventPromises);
      eventArrays.forEach(events => allEvents.push(...events));

      // Sort events by start date
      allEvents.sort((a, b) => a.start.getTime() - b.start.getTime());

      // Limit results
      const limitedEvents = allEvents.slice(0, maxEvents);

      // Cache aggregated results
      this.cacheService.setCachedAggregatedEvents(sourceIds, limitedEvents, AppConstants.CACHE_DURATION_MINUTES);

      // Log any errors
      if (errors.length > 0) {
        console.warn('Some calendar sources failed to load:', errors);
      }

      return limitedEvents;
    } catch (error) {
      console.error('Error getting events from multiple sources:', error);
      throw error;
    }
  }

  /**
   * Search events across calendar sources
   */
  public async searchEvents(sources: ICalendarSource[], query: string, maxResults: number = 50): Promise<ICalendarEvent[]> {
    try {
      // Validate search query
      const queryValidation = ValidationUtils.validateSearchQuery(query);
      if (!queryValidation.isValid) {
        throw new Error(`Invalid search query: ${queryValidation.errors.map(e => e.message).join(', ')}`);
      }

      // Check cache first
      const cachedResults = this.cacheService.getCachedSearchResults(query);
      if (cachedResults) {
        return cachedResults.slice(0, maxResults);
      }

      const allResults: ICalendarEvent[] = [];
      const errors: string[] = [];

      // Search SharePoint calendars
      const sharePointSources = sources.filter(s => s.type === CalendarSourceType.SharePoint);
      if (sharePointSources.length > 0) {
        try {
          const sharePointResults = await this.sharePointService.searchEvents(sharePointSources, query, maxResults);
          allResults.push(...sharePointResults);
        } catch (error) {
          console.error('SharePoint search failed:', error);
          errors.push('SharePoint search failed');
        }
      }

      // Search Exchange calendars
      const exchangeSources = sources.filter(s => s.type === CalendarSourceType.Exchange);
      if (exchangeSources.length > 0) {
        try {
          const exchangeResults = await this.exchangeService.searchEvents(exchangeSources, query, maxResults);
          allResults.push(...exchangeResults);
        } catch (error) {
          console.error('Exchange search failed:', error);
          errors.push('Exchange search failed');
        }
      }

      // Remove duplicates and sort by relevance/date
      const uniqueResults = this.removeDuplicateEvents(allResults);
      uniqueResults.sort((a, b) => {
        // Sort by relevance (title match first, then by date)
        const aRelevance = a.title.toLowerCase().includes(query.toLowerCase()) ? 1 : 0;
        const bRelevance = b.title.toLowerCase().includes(query.toLowerCase()) ? 1 : 0;
        
        if (aRelevance !== bRelevance) {
          return bRelevance - aRelevance; // Higher relevance first
        }
        
        return a.start.getTime() - b.start.getTime(); // Earlier events first
      });

      const limitedResults = uniqueResults.slice(0, maxResults);

      // Cache search results
      this.cacheService.setCachedSearchResults(query, limitedResults, 5); // Shorter cache for search

      return limitedResults;
    } catch (error) {
      console.error('Error searching events:', error);
      throw error;
    }
  }

  /**
   * Get events for specific date range
   */
  public async getEventsForDateRange(
    sources: ICalendarSource[], 
    startDate: Date, 
    endDate: Date, 
    maxEvents: number = 1000
  ): Promise<ICalendarEvent[]> {
    try {
      // Validate date range
      const dateValidation = ValidationUtils.validateDateRange(startDate, endDate);
      if (!dateValidation.isValid) {
        throw new Error(`Invalid date range: ${dateValidation.errors.map(e => e.message).join(', ')}`);
      }

      // Check cache first
      const sourceIds = sources.map(s => s.id);
      const cachedEvents = this.cacheService.getCachedDateRangeEvents(startDate, endDate, sourceIds);
      if (cachedEvents) {
        return cachedEvents.slice(0, maxEvents);
      }

      const allEvents: ICalendarEvent[] = [];
      const errors: string[] = [];

      // Get events from SharePoint sources
      const sharePointSources = sources.filter(s => s.type === CalendarSourceType.SharePoint);
      if (sharePointSources.length > 0) {
        try {
          const sharePointEvents = await this.sharePointService.getEventsForDateRange(
            sharePointSources, startDate, endDate, maxEvents
          );
          allEvents.push(...sharePointEvents);
        } catch (error) {
          console.error('SharePoint date range query failed:', error);
          errors.push('SharePoint query failed');
        }
      }

      // Get events from Exchange sources
      const exchangeSources = sources.filter(s => s.type === CalendarSourceType.Exchange);
      if (exchangeSources.length > 0) {
        try {
          const exchangeEvents = await this.exchangeService.getEventsForDateRange(
            exchangeSources, startDate, endDate, maxEvents
          );
          allEvents.push(...exchangeEvents);
        } catch (error) {
          console.error('Exchange date range query failed:', error);
          errors.push('Exchange query failed');
        }
      }

      // Remove duplicates and sort
      const uniqueEvents = this.removeDuplicateEvents(allEvents);
      uniqueEvents.sort((a, b) => a.start.getTime() - b.start.getTime());

      const limitedEvents = uniqueEvents.slice(0, maxEvents);

      // Cache results
      this.cacheService.setCachedDateRangeEvents(startDate, endDate, sourceIds, limitedEvents, AppConstants.CACHE_DURATION_MINUTES);

      return limitedEvents;
    } catch (error) {
      console.error('Error getting events for date range:', error);
      throw error;
    }
  }
private convertToExtendedCalendarEvent(event: ICalendarEvent): IExtendedCalendarEvent {
  // Convert attendees to include index signature
  const convertedAttendees = event.attendees?.map(attendee => ({
    name: attendee.name,
    email: attendee.email,
    response: attendee.response,
    type: attendee.type,
    isOrganizer: false, // Default value for extended interface
    // Add index signature compatibility
    [Symbol.toPrimitive]: () => `${attendee.name} <${attendee.email}>` // Example implementation
  } as IEventAttendee)) || [];

  return {
    id: event.id,
    title: event.title,
    description: event.description,
    start: event.start,
    end: event.end,
    location: event.location,
    category: event.category,
    isAllDay: event.isAllDay,
    isRecurring: event.isRecurring,
    calendarId: event.calendarId,
    calendarTitle: event.calendarTitle,
    calendarType: event.calendarType,
    organizer: event.organizer,
    created: event.created,
    modified: event.modified,
    webUrl: event.webUrl,
    color: event.color,
    importance: event.importance,
    sensitivity: event.sensitivity,
    showAs: event.showAs,
    attendees: convertedAttendees,
    attachments: event.attachments?.map(attachment => ({
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      url: attachment.url,
      isInline: attachment.isInline,
      // Add index signature compatibility
      id: undefined,
      content: undefined,
      lastModified: undefined
    } as IEventAttachment)) || [],
    // Extended properties with default values
    timeZone: 'UTC',
    reminderMinutes: [],
    isPrivate: false,
    isCancelled: false,
    responseStatus: 'none',
    meetingType: 'standard',
    onlineMeetingUrl: undefined,
    recurrencePattern: undefined,
    exceptions: [],
    masterSeriesId: event.masterSeriesId,
    isException: event.isException,
    tags: event.tags || [],
    customProperties: {}
  };
}
  /**
   * Advanced search with criteria
   */
public async advancedSearch(criteria: IEventSearchCriteria): Promise<IEventSearchResult> {
  const startTime = Date.now();
  
  try {
    // Get all events first (this could be optimized to filter at source level)
    const sources = await this.getCalendarSources(true);
    const filteredSources = criteria.calendarIds 
      ? sources.filter(s => criteria.calendarIds!.includes(s.id))
      : sources;

    let events: ICalendarEvent[] = [];

    if (criteria.startDate && criteria.endDate) {
      events = await this.getEventsForDateRange(filteredSources, criteria.startDate, criteria.endDate);
    } else {
      events = await this.getEventsFromSources(filteredSources);
    }

    // Apply additional filters
    let filteredEvents = this.applyAdvancedFilters(events, criteria);

    // Apply search query if provided
    if (criteria.query) {
      filteredEvents = this.filterEventsByQuery(filteredEvents, criteria.query);
    }

    // Convert ICalendarEvent[] to IExtendedCalendarEvent[]
    const extendedEvents: IExtendedCalendarEvent[] = filteredEvents.map(event => 
      this.convertToExtendedCalendarEvent(event)
    );

    const searchDuration = Date.now() - startTime;

    return {
      events: extendedEvents,
      totalCount: extendedEvents.length,
      searchDuration,
      hasMore: false, // Could implement pagination here
      nextPageToken: undefined
    };
  } catch (error) {
    console.error('Advanced search failed:', error);
    throw error;
  }
}


  /**
   * Create a new event (if supported by calendar type)
   */
  public async createEvent?(calendarId: string, eventData: IEventCreateRequest): Promise<ICalendarEvent> {
    try {
      // Validate event data
      const validation = ValidationUtils.validateEventCreateRequest(eventData);
      if (!validation.isValid) {
        throw new Error(`Invalid event data: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Find the calendar source
      const sources = await this.getCalendarSources(true);
      const source = sources.find(s => s.id === calendarId);
      if (!source) {
        throw new Error('Calendar not found');
      }

      let createdEvent: ICalendarEvent | null = null;

      // Route to appropriate service
      switch (source.type) {
        case CalendarSourceType.Exchange: {
          // Convert IEventCreateRequest to Record<string, unknown>
          const exchangeEventData: Record<string, unknown> = {
            subject: eventData.title,
            body: {
              content: eventData.description || '',
              contentType: 'html'
            },
            start: {
              dateTime: eventData.start.toISOString(),
              timeZone: 'UTC'
            },
            end: {
              dateTime: eventData.end.toISOString(),
              timeZone: 'UTC'
            },
            location: eventData.location ? {
              displayName: eventData.location
            } : undefined,
            categories: eventData.category ? [eventData.category] : undefined,
            isAllDay: eventData.isAllDay || false,
            importance: eventData.importance || 'normal',
            sensitivity: eventData.sensitivity || 'normal',
            showAs: eventData.showAs || 'busy',
            attendees: eventData.attendees?.map(attendee => ({
              emailAddress: {
                address: attendee.email,
                name: attendee.name
              },
              type: attendee.type || 'required'
            }))
          };
          
          createdEvent = await this.exchangeService.createEvent(calendarId, exchangeEventData) || null;
          break;
        }
        case CalendarSourceType.SharePoint: {
          throw new Error('SharePoint event creation not implemented');
        }
        default: {
          throw new Error(`Event creation not supported for calendar type: ${source.type}`);
        }
      }

      if (!createdEvent) {
        throw new Error('Failed to create event');
      }

      // Invalidate cache for this calendar
      this.cacheService.clearCalendarSpecificCache(calendarId);

      return createdEvent;
    } catch (error) {
      console.error('Error creating event:', error);
      throw error;
    }
  }

  /**
   * Update an existing event (if supported by calendar type)
   */
  public async updateEvent?(calendarId: string, eventId: string, eventData: IEventUpdateRequest): Promise<ICalendarEvent> {
    try {
      // Validate event data
      const validation = ValidationUtils.validateEventUpdateRequest(eventData);
      if (!validation.isValid) {
        throw new Error(`Invalid event data: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Find the calendar source
      const sources = await this.getCalendarSources(true);
      const source = sources.find(s => s.id === calendarId);
      if (!source) {
        throw new Error('Calendar not found');
      }

      let updatedEvent: ICalendarEvent | null = null;

      // Route to appropriate service
      switch (source.type) {
        case CalendarSourceType.Exchange: {
          // Convert IEventUpdateRequest to Record<string, unknown>
          const exchangeUpdateData: Record<string, unknown> = {};
          
          if (eventData.title !== undefined) {
            exchangeUpdateData.subject = eventData.title;
          }
          if (eventData.description !== undefined) {
            exchangeUpdateData.body = {
              content: eventData.description,
              contentType: 'html'
            };
          }
          if (eventData.start !== undefined) {
            exchangeUpdateData.start = {
              dateTime: eventData.start.toISOString(),
              timeZone: 'UTC'
            };
          }
          if (eventData.end !== undefined) {
            exchangeUpdateData.end = {
              dateTime: eventData.end.toISOString(),
              timeZone: 'UTC'
            };
          }
          if (eventData.location !== undefined) {
            exchangeUpdateData.location = eventData.location ? {
              displayName: eventData.location
            } : null;
          }
          if (eventData.category !== undefined) {
            exchangeUpdateData.categories = eventData.category ? [eventData.category] : [];
          }
          if (eventData.isAllDay !== undefined) {
            exchangeUpdateData.isAllDay = eventData.isAllDay;
          }
          if (eventData.importance !== undefined) {
            exchangeUpdateData.importance = eventData.importance;
          }
          if (eventData.sensitivity !== undefined) {
            exchangeUpdateData.sensitivity = eventData.sensitivity;
          }
          if (eventData.showAs !== undefined) {
            exchangeUpdateData.showAs = eventData.showAs;
          }
          if (eventData.attendees !== undefined) {
            exchangeUpdateData.attendees = eventData.attendees.map(attendee => ({
              emailAddress: {
                address: attendee.email,
                name: attendee.name
              },
              type: attendee.type || 'required'
            }));
          }
          
          updatedEvent = await this.exchangeService.updateEvent(calendarId, eventId, exchangeUpdateData) || null;
          break;
        }
        case CalendarSourceType.SharePoint: {
          throw new Error('SharePoint event update not implemented');
        }
        default: {
          throw new Error(`Event update not supported for calendar type: ${source.type}`);
        }
      }

      if (!updatedEvent) {
        throw new Error('Failed to update event');
      }

      // Invalidate cache for this calendar
      this.cacheService.clearCalendarSpecificCache(calendarId);

      return updatedEvent;
    } catch (error) {
      console.error('Error updating event:', error);
      throw error;
    }
  }

  /**
   * Delete an event (if supported by calendar type)
   */
  public async deleteEvent?(calendarId: string, eventId: string): Promise<boolean> {
    try {
      // Find the calendar source
      const sources = await this.getCalendarSources(true);
      const source = sources.find(s => s.id === calendarId);
      if (!source) {
        throw new Error('Calendar not found');
      }

      let deleted = false;

      // Route to appropriate service
      switch (source.type) {
        case CalendarSourceType.Exchange: {
          deleted = await this.exchangeService.deleteEvent(calendarId, eventId);
          break;
        }
        case CalendarSourceType.SharePoint: {
          throw new Error('SharePoint event deletion not implemented');
        }
        default: {
          throw new Error(`Event deletion not supported for calendar type: ${source.type}`);
        }
      }

      if (deleted) {
        // Invalidate cache for this calendar
        this.cacheService.clearCalendarSpecificCache(calendarId);
      }

      return deleted;
    } catch (error) {
      console.error('Error deleting event:', error);
      throw error;
    }
  }

  /**
   * Get calendar statistics
   */
  public async getCalendarStatistics(sources: ICalendarSource[]): Promise<{ [calendarId: string]: Record<string, unknown> }> {
    const statistics: { [calendarId: string]: Record<string, unknown> } = {};

    for (const source of sources) {
      try {
        switch (source.type) {
          case CalendarSourceType.SharePoint: {
            statistics[source.id] = await this.sharePointService.getCalendarStatistics(source);
            break;
          }
          case CalendarSourceType.Exchange: {
            statistics[source.id] = await this.exchangeService.getCalendarStatistics(source.id);
            break;
          }
        }
      } catch (error) {
        console.warn(`Failed to get statistics for ${source.title}:`, error);
        statistics[source.id] = { error: 'Failed to load statistics' };
      }
    }

    return statistics;
  }

  /**
   * Get user's free/busy information
   */
  public async getFreeBusyInfo(
    calendarIds: string[], 
    startDate: Date, 
    endDate: Date
  ): Promise<{ [calendarId: string]: Record<string, unknown> }> {
    const freeBusyData: { [calendarId: string]: Record<string, unknown> } = {};

    try {
      // Currently only supported for Exchange calendars
      const freeBusyInfo = await this.exchangeService.getFreeBusyInfo(calendarIds, startDate, endDate);
      if (freeBusyInfo) {
        freeBusyInfo.forEach((info: Record<string, unknown>, index: number) => {
          freeBusyData[calendarIds[index]] = info;
        });
      }
    } catch (error) {
      console.error('Error getting free/busy information:', error);
    }

    return freeBusyData;
  }

  /**
   * Get working hours for the current user
   */
  public async getUserWorkingHours(): Promise<Record<string, unknown> | undefined> {
    try {
      return await this.exchangeService.getUserWorkingHours();
    } catch (error) {
      console.error('Error getting user working hours:', error);
      return undefined;
    }
  }

  /**
   * Get user's timezone
   */
  public async getUserTimeZone(): Promise<string> {
    try {
      return await this.exchangeService.getUserTimeZone();
    } catch (error) {
      console.error('Error getting user timezone:', error);
      return 'UTC';
    }
  }

  /**
   * Check Graph API access
   */
  public async checkGraphAccess(): Promise<boolean> {
    try {
      return await this.exchangeService.checkGraphAccess();
    } catch (error) {
      console.error('Error checking Graph access:', error);
      return false;
    }
  }

  /**
   * Get event conflicts for a time period
   */
  public async getEventConflicts(
    sources: ICalendarSource[], 
    startDate: Date, 
    endDate: Date
  ): Promise<{ event1: ICalendarEvent; event2: ICalendarEvent; overlapMinutes: number }[]> {
    try {
      const events = await this.getEventsForDateRange(sources, startDate, endDate);
      const conflicts: { event1: ICalendarEvent; event2: ICalendarEvent; overlapMinutes: number }[] = [];

      // Find overlapping events
      for (let i = 0; i < events.length; i++) {
        for (let j = i + 1; j < events.length; j++) {
          const event1 = events[i];
          const event2 = events[j];

          // Skip all-day events for conflict detection
          if (event1.isAllDay || event2.isAllDay) continue;

          // Check for overlap
          const overlap = this.calculateEventOverlap(event1, event2);
          if (overlap > 0) {
            conflicts.push({
              event1,
              event2,
              overlapMinutes: overlap
            });
          }
        }
      }

      return conflicts;
    } catch (error) {
      console.error('Error getting event conflicts:', error);
      return [];
    }
  }

  /**
   * Clear all cached data
   */
  public clearCache(): void {
    this.cacheService.clearCalendarCache();
  }

  /**
   * Clear cache for specific calendar
   */
  public clearCalendarCache(calendarId: string): void {
    this.cacheService.clearCalendarSpecificCache(calendarId);
  }

  /**
   * Get cache health information
   */
  public getCacheHealth(): Record<string, unknown> {
    return this.cacheService.getCacheHealth();
  }

  /**
   * Optimize cache performance
   */
  public optimizeCache(): void {
    this.cacheService.optimizeCache();
  }

  /**
   * Export cache for debugging
   */
  public exportCache(): Record<string, unknown> {
    return this.cacheService.exportCache();
  }

  // Private helper methods

  /**
   * Validate and enhance calendar sources
   */
  private async validateAndEnhanceSources(sources: ICalendarSource[]): Promise<ICalendarSource[]> {
    const validatedSources: ICalendarSource[] = [];

    for (const source of sources) {
      try {
        // Validate source
        const validation = ValidationUtils.validateCalendarSource(source);
        if (validation.isValid) {
          // Enhance source with additional metadata
          const enhancedSource = {
            ...source,
            color: source.color || ColorUtils.generateColorFromString(source.title),
            isEnabled: source.isEnabled !== false // Default to true if not specified
          };
          validatedSources.push(enhancedSource);
        } else {
          console.warn(`Invalid calendar source ${source.title}:`, validation.errors);
        }
      } catch (error) {
        console.warn(`Error validating source ${source.title}:`, error);
      }
    }

    return validatedSources;
  }

  /**
   * Validate and enhance events
   */
  private async validateAndEnhanceEvents(events: ICalendarEvent[], source: ICalendarSource): Promise<ICalendarEvent[]> {
    const validatedEvents: ICalendarEvent[] = [];

    for (const event of events) {
      try {
        // Validate event
        const validation = ValidationUtils.validateCalendarEvent(event);
        if (validation.isValid) {
          // Enhance event with additional metadata
          const enhancedEvent = {
            ...event,
            color: event.color || source.color || ColorUtils.generateColorFromString(source.title),
            calendarTitle: source.title,
            calendarType: source.type
          };
          validatedEvents.push(enhancedEvent);
        } else {
          console.warn(`Invalid event ${event.title}:`, validation.errors);
        }
      } catch (error) {
        console.warn(`Error validating event ${event.title}:`, error);
      }
    }

    return validatedEvents;
  }

  /**
   * Remove duplicate events based on title, start time, and organizer
   */
  private removeDuplicateEvents(events: ICalendarEvent[]): ICalendarEvent[] {
    const seen = new Set<string>();
    const uniqueEvents: ICalendarEvent[] = [];

    for (const event of events) {
      const key = `${event.title}_${event.start.getTime()}_${event.organizer}`.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        uniqueEvents.push(event);
      }
    }

    return uniqueEvents;
  }

  /**
   * Apply advanced filters to events
   */
  private applyAdvancedFilters(events: ICalendarEvent[], criteria: IEventSearchCriteria): ICalendarEvent[] {
    let filtered = events;

    // Filter by categories
    if (criteria.categories && criteria.categories.length > 0) {
      filtered = filtered.filter(event => 
        event.category && criteria.categories!.includes(event.category)
      );
    }

    // Filter by organizers
    if (criteria.organizers && criteria.organizers.length > 0) {
      filtered = filtered.filter(event => 
        criteria.organizers!.some(org => 
          event.organizer.toLowerCase().includes(org.toLowerCase())
        )
      );
    }

    // Filter by attendees
    if (criteria.attendees && criteria.attendees.length > 0) {
      filtered = filtered.filter(event => 
        event.attendees && event.attendees.some(attendee =>
          criteria.attendees!.some(searchAttendee =>
            attendee.email.toLowerCase().includes(searchAttendee.toLowerCase()) ||
            attendee.name.toLowerCase().includes(searchAttendee.toLowerCase())
          )
        )
      );
    }

    // Filter by importance
    if (criteria.importance && criteria.importance.length > 0) {
      filtered = filtered.filter(event => 
        event.importance && criteria.importance!.includes(event.importance)
      );
    }

    // Filter by sensitivity
    if (criteria.sensitivity && criteria.sensitivity.length > 0) {
      filtered = filtered.filter(event => 
        event.sensitivity && criteria.sensitivity!.includes(event.sensitivity)
      );
    }

    // Filter by attachments
    if (criteria.hasAttachments !== undefined) {
      filtered = filtered.filter(event => 
        criteria.hasAttachments ? 
        (event.attachments && event.attachments.length > 0) : 
        (!event.attachments || event.attachments.length === 0)
      );
    }

    // Filter by recurring status
    if (criteria.isRecurring !== undefined) {
      filtered = filtered.filter(event => event.isRecurring === criteria.isRecurring);
    }

    // Filter by all-day status
    if (criteria.isAllDay !== undefined) {
      filtered = filtered.filter(event => event.isAllDay === criteria.isAllDay);
    }

    // Filter by tags
    if (criteria.tags && criteria.tags.length > 0) {
      filtered = filtered.filter(event => 
        event.tags && event.tags.some(tag => 
          criteria.tags!.includes(tag)
        )
      );
    }

    return filtered;
  }

  /**
   * Filter events by search query
   */
  private filterEventsByQuery(events: ICalendarEvent[], query: string): ICalendarEvent[] {
    const searchTerms = query.toLowerCase().split(' ').filter(term => term.length > 0);
    
    return events.filter(event => {
      const searchableText = [
        event.title,
        event.description || '',
        event.location || '',
        event.organizer,
        event.category || '',
        ...(event.attendees?.map(a => `${a.name} ${a.email}`) || [])
      ].join(' ').toLowerCase();

      return searchTerms.every(term => searchableText.includes(term));
    });
  }

  /**
   * Calculate overlap between two events in minutes
   */
  private calculateEventOverlap(event1: ICalendarEvent, event2: ICalendarEvent): number {
    const start1 = event1.start.getTime();
    const end1 = event1.end.getTime();
    const start2 = event2.start.getTime();
    const end2 = event2.end.getTime();

    // No overlap if one event ends before the other starts
    if (end1 <= start2 || end2 <= start1) {
      return 0;
    }

    // Calculate overlap
    const overlapStart = Math.max(start1, start2);
    const overlapEnd = Math.min(end1, end2);
    const overlapMs = overlapEnd - overlapStart;

    return Math.max(0, Math.floor(overlapMs / (1000 * 60))); // Convert to minutes
  }

  /**
   * Get service health status
   */
  public async getServiceHealth(): Promise<{
    sharePoint: { available: boolean; responseTime?: number; error?: string };
    exchange: { available: boolean; responseTime?: number; error?: string };
    cache: { healthy: boolean; size: number; hitRate?: number };
  }> {
    const health = {
      sharePoint: { available: false, responseTime: 0, error: undefined as string | undefined },
      exchange: { available: false, responseTime: 0, error: undefined as string | undefined },
      cache: { healthy: false, size: 0, hitRate: undefined as number | undefined }
    };

    // Test SharePoint connectivity
    try {
      const start = Date.now();
      await this.sharePointService.getSharePointCalendars();
      health.sharePoint.available = true;
      health.sharePoint.responseTime = Date.now() - start;
    } catch (error) {
      health.sharePoint.error = error instanceof Error ? error.message : 'Unknown error';
    }

    // Test Exchange connectivity
    try {
      const start = Date.now();
      health.exchange.available = await this.exchangeService.checkGraphAccess();
      health.exchange.responseTime = Date.now() - start;
    } catch (error) {
      health.exchange.error = error instanceof Error ? error.message : 'Unknown error';
    }

    // Check cache health
    try {
      const cacheHealth = this.cacheService.getCacheHealth();
      health.cache = {
        healthy: cacheHealth.expiredItems < cacheHealth.totalItems * 0.1, // Less than 10% expired
        size: cacheHealth.totalItems,
        hitRate: cacheHealth.hitRate || 0
      };
    } catch {
      health.cache = { healthy: false, size: 0, hitRate: 0 };
    }

    return health;
  }

  /**
   * Get analytics data for calendar usage
   */
  public async getAnalytics(
    sources: ICalendarSource[], 
    startDate: Date, 
    endDate: Date
  ): Promise<{
    totalEvents: number;
    eventsByCalendar: { [calendarId: string]: number };
    eventsByCategory: { [category: string]: number };
    eventsByMonth: { [month: string]: number };
    averageEventDuration: number;
    topOrganizers: { name: string; count: number }[];
    busyHours: { hour: number; count: number }[];
    recurringVsOneTime: { recurring: number; oneTime: number };
  }> {
    try {
      const events = await this.getEventsForDateRange(sources, startDate, endDate);
      
      const analytics = {
        totalEvents: events.length,
        eventsByCalendar: {} as { [calendarId: string]: number },
        eventsByCategory: {} as { [category: string]: number },
        eventsByMonth: {} as { [month: string]: number },
        averageEventDuration: 0,
        topOrganizers: [] as { name: string; count: number }[],
        busyHours: [] as { hour: number; count: number }[],
        recurringVsOneTime: { recurring: 0, oneTime: 0 }
      };

      if (events.length === 0) return analytics;

      // Calculate metrics
      let totalDuration = 0;
      const organizerCounts: { [name: string]: number } = {};
      const hourCounts: { [hour: number]: number } = {};

      events.forEach(event => {
        // Events by calendar
        analytics.eventsByCalendar[event.calendarId] = (analytics.eventsByCalendar[event.calendarId] || 0) + 1;

        // Events by category
        const category = event.category || 'Uncategorized';
        analytics.eventsByCategory[category] = (analytics.eventsByCategory[category] || 0) + 1;

        // Events by month
        const monthKey = DateUtils.formatDate(event.start, 'YYYY-MM');
        analytics.eventsByMonth[monthKey] = (analytics.eventsByMonth[monthKey] || 0) + 1;

        // Duration calculation
        if (!event.isAllDay) {
          const duration = event.end.getTime() - event.start.getTime();
          totalDuration += duration;
        }

        // Organizer counts
        organizerCounts[event.organizer] = (organizerCounts[event.organizer] || 0) + 1;

        // Busy hours (for non-all-day events)
        if (!event.isAllDay) {
          const startHour = event.start.getHours();
          hourCounts[startHour] = (hourCounts[startHour] || 0) + 1;
        }

        // Recurring vs one-time
        if (event.isRecurring) {
          analytics.recurringVsOneTime.recurring++;
        } else {
          analytics.recurringVsOneTime.oneTime++;
        }
      });

      // Calculate average duration (in minutes)
      const nonAllDayEvents = events.filter(e => !e.isAllDay);
      analytics.averageEventDuration = nonAllDayEvents.length > 0 
        ? Math.round(totalDuration / nonAllDayEvents.length / (1000 * 60))
        : 0;

      // Top organizers
      analytics.topOrganizers = Object.entries(organizerCounts)
        .map(([name, count]) => ({ name, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

      // Busy hours
      analytics.busyHours = Object.entries(hourCounts)
        .map(([hour, count]) => ({ hour: parseInt(hour), count }))
        .sort((a, b) => b.count - a.count);

      return analytics;
    } catch (error) {
      console.error('Error generating analytics:', error);
      throw error;
    }
  }

  /**
   * Get event recommendations based on patterns
   */
  public async getEventRecommendations(
    sources: ICalendarSource[],
    targetDate: Date
  ): Promise<{
    suggestedMeetings: { title: string; reason: string; confidence: number }[];
    availableSlots: { start: Date; end: Date; duration: number }[];
    conflictWarnings: { message: string; severity: 'low' | 'medium' | 'high' }[];
  }> {
    try {
      // Get events for context (past month and next week)
      const pastMonth = DateUtils.subtractTime(targetDate, 30, 'days');
      const nextWeek = DateUtils.addTime(targetDate, 7, 'days');
      const events = await this.getEventsForDateRange(sources, pastMonth, nextWeek);

      const recommendations = {
        suggestedMeetings: [] as { title: string; reason: string; confidence: number }[],
        availableSlots: [] as { start: Date; end: Date; duration: number }[],
        conflictWarnings: [] as { message: string; severity: 'low' | 'medium' | 'high' }[]
      };

      // Analyze patterns for suggestions
      const recurringPatterns = this.analyzeRecurringPatterns(events);

      // Generate meeting suggestions based on patterns
      recurringPatterns.forEach(pattern => {
        if (pattern.confidence > 0.7) {
          recommendations.suggestedMeetings.push({
            title: pattern.title,
            reason: `Recurring ${pattern.frequency} meeting pattern detected`,
            confidence: pattern.confidence
          });
        }
      });

      // Find available time slots
      const workingHours = await this.getUserWorkingHours();
      recommendations.availableSlots = this.findAvailableSlots(
        events,
        targetDate,
        workingHours || { startTime: '09:00', endTime: '17:00', daysOfWeek: [1, 2, 3, 4, 5] }
      );

      // Generate conflict warnings
      const conflicts = await this.getEventConflicts(sources, targetDate, DateUtils.addTime(targetDate, 1, 'days'));
      conflicts.forEach(conflict => {
        recommendations.conflictWarnings.push({
          message: `Potential conflict between "${conflict.event1.title}" and "${conflict.event2.title}"`,
          severity: conflict.overlapMinutes > 60 ? 'high' : conflict.overlapMinutes > 30 ? 'medium' : 'low'
        });
      });

      return recommendations;
    } catch (error) {
      console.error('Error generating recommendations:', error);
      return {
        suggestedMeetings: [],
        availableSlots: [],
        conflictWarnings: []
      };
    }
  }

  /**
   * Bulk operations for events
   */
  public async bulkUpdateEvents(
    operations: Array<{
      type: 'create' | 'update' | 'delete';
      calendarId: string;
      eventId?: string;
      eventData?: IEventCreateRequest | IEventUpdateRequest;
    }>
  ): Promise<{
    successful: number;
    failed: number;
    errors: Array<{ operation: Record<string, unknown>; error: string }>;
  }> {
    const results = {
      successful: 0,
      failed: 0,
      errors: [] as Array<{ operation: Record<string, unknown>; error: string }>
    };

    for (const operation of operations) {
      try {
        switch (operation.type) {
          case 'create': {
            if (operation.eventData) {
              await this.createEvent!(operation.calendarId, operation.eventData as IEventCreateRequest);
              results.successful++;
            }
            break;
          }
          case 'update': {
            if (operation.eventId && operation.eventData) {
              await this.updateEvent!(operation.calendarId, operation.eventId, operation.eventData as IEventUpdateRequest);
              results.successful++;
            }
            break;
          }
          case 'delete': {
            if (operation.eventId) {
              await this.deleteEvent!(operation.calendarId, operation.eventId);
              results.successful++;
            }
            break;
          }
        }
      } catch (error) {
        results.failed++;
        results.errors.push({
          operation,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    return results;
  }

  /**
   * Export events in various formats
   */
  public async exportEvents(
    sources: ICalendarSource[],
    startDate: Date,
    endDate: Date,
    format: 'ics' | 'csv' | 'json'
  ): Promise<{ data: string; filename: string; mimeType: string }> {
    try {
      const events = await this.getEventsForDateRange(sources, startDate, endDate);
      const dateRange = `${DateUtils.formatDate(startDate, 'YYYY-MM-DD')}_to_${DateUtils.formatDate(endDate, 'YYYY-MM-DD')}`;

      switch (format) {
        case 'ics': {
          // Use the public exportToICS method and capture the content
          const icsContent = this.generateICSContent(events);
          return {
            data: icsContent,
            filename: `calendar_events_${dateRange}.ics`,
            mimeType: 'text/calendar'
          };
        }
        case 'csv': {
          const csvContent = this.generateCSVContent(events);
          return {
            data: csvContent,
            filename: `calendar_events_${dateRange}.csv`,
            mimeType: 'text/csv'
          };
        }
        case 'json': {
          const jsonContent = JSON.stringify({
            exportDate: new Date().toISOString(),
            dateRange: { start: startDate, end: endDate },
            totalEvents: events.length,
            events: events
          }, null, 2);
          return {
            data: jsonContent,
            filename: `calendar_events_${dateRange}.json`,
            mimeType: 'application/json'
          };
        }
        default: {
          throw new Error(`Unsupported export format: ${format}`);
        }
      }
    } catch (error) {
      console.error('Error exporting events:', error);
      throw error;
    }
  }

  // ... Rest of the private helper methods continue here

  private generateICSContent(events: ICalendarEvent[]): string {
    // Implementation for ICS generation
    return 'ICS content generation not implemented in this snippet';
  }

  private generateCSVContent(events: ICalendarEvent[]): string {
    // Implementation for CSV generation
    return 'CSV content generation not implemented in this snippet';
  }

  private analyzeRecurringPatterns(events: ICalendarEvent[]): Array<{
    title: string;
    frequency: string;
    confidence: number;
  }> {
    // Implementation for pattern analysis
    return [];
  }

  private findAvailableSlots(
    events: ICalendarEvent[],
    targetDate: Date,
    workingHours: Record<string, unknown>
  ): Array<{ start: Date; end: Date; duration: number }> {
    // Implementation for finding available slots
    return [];
  }
}