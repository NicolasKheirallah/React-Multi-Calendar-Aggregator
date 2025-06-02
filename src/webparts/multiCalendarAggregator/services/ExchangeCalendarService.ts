import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { ICalendarEvent, ICalendarSource, CalendarSourceType } from '../models/ICalendarModels';
import { AppConstants } from '../constants/AppConstants';
import { ColorUtils } from '../utils/ColorUtils';

interface GraphEvent {
  id: string;
  subject: string;
  body?: {
    content: string;
  };
  start: {
    dateTime: string;
    timeZone?: string;
  };
  end: {
    dateTime: string;
    timeZone?: string;
  };
  location?: {
    displayName: string;
  };
  categories?: string[];
  isAllDay: boolean;
  recurrence?: unknown;
  organizer?: {
    emailAddress?: {
      name: string;
    };
  };
  createdDateTime: string;
  lastModifiedDateTime: string;
  webLink: string;
  importance?: string;
  sensitivity?: string;
  showAs?: string;
  attendees?: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
    status: {
      response: string;
    };
    type: string;
  }>;
}

interface GraphCalendar {
  id: string;
  name: string;
  color?: string;
  owner?: {
    name: string;
  };
  canEdit?: boolean;
  canShare?: boolean;
}


export class ExchangeCalendarService {
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | undefined;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Initialize Microsoft Graph client
   */
  private async initializeGraphClient(): Promise<void> {
    if (!this.graphClient) {
      try {
        this.graphClient = await this.context.msGraphClientFactory.getClient('3');
      } catch (error) {
        console.error('Failed to initialize Graph client:', error);
        throw new Error(AppConstants.ERROR_MESSAGES.GRAPH_API_ERROR);
      }
    }
  }

  /**
   * Get Exchange Online calendars
   */
  public async getExchangeCalendars(): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        throw new Error('Graph client initialization failed');
      }

      // Get user's calendars
      const userCalendars = await this.getUserCalendars();
      calendars.push(...userCalendars);

      // Get shared calendars
      const sharedCalendars = await this.getSharedCalendars();
      calendars.push(...sharedCalendars);

      // Get group calendars
      const groupCalendars = await this.getGroupCalendars();
      calendars.push(...groupCalendars);

    } catch (error) {
      console.error('Error getting Exchange calendars:', error);
      // Don't throw - just log and continue without Exchange calendars
    }

    return calendars;
  }

  /**
   * Get user's personal calendars
   */
  private async getUserCalendars(): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      if (!this.graphClient) return calendars;

      const response = await this.graphClient
        .api('/me/calendars')
        .select(AppConstants.GRAPH_CALENDAR_FIELDS)
        .get();

      for (const calendar of response.value || []) {
        calendars.push({
          id: calendar.id,
          title: calendar.name,
          description: `Personal calendar${calendar.owner?.name ? ` (${calendar.owner.name})` : ''}`,
          type: CalendarSourceType.Exchange,
          url: 'https://outlook.office365.com/calendar/view/month',
          siteTitle: 'Exchange Online',
          siteUrl: 'https://outlook.office365.com',
          color: ColorUtils.mapExchangeColor(calendar.color),
          isEnabled: true,
          canEdit: calendar.canEdit || false,
          canShare: calendar.canShare || false
        });
      }
    } catch (error) {
      console.error('Error getting user calendars:', error);
    }

    return calendars;
  }

  /**
   * Get shared calendars
   */
  private async getSharedCalendars(): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      if (!this.graphClient) return calendars;

      const response = await this.graphClient
        .api('/me/calendarGroups')
        .expand('calendars')
        .get();

      for (const group of response.value || []) {
        if (group.name !== 'My Calendars' && group.calendars) {
          for (const calendar of group.calendars) {
            calendars.push({
              id: calendar.id,
              title: `${calendar.name} (${group.name})`,
              description: `Shared calendar from ${group.name}`,
              type: CalendarSourceType.Exchange,
              url: 'https://outlook.office365.com/calendar/view/month',
              siteTitle: 'Exchange Online',
              siteUrl: 'https://outlook.office365.com',
              color: ColorUtils.mapExchangeColor(calendar.color),
              isEnabled: true,
              canEdit: calendar.canEdit || false,
              canShare: calendar.canShare || false
            });
          }
        }
      }
    } catch (error) {
      console.error('Error getting shared calendars:', error);
    }

    return calendars;
  }

  /**
   * Get group calendars (if user has access)
   */
  private async getGroupCalendars(): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      if (!this.graphClient) return calendars;

      // Get groups user is member of
      const groupsResponse = await this.graphClient
        .api('/me/memberOf')
        .filter("@odata.type eq 'microsoft.graph.group'")
        .select('id,displayName,mail')
        .top(50)
        .get();

      for (const group of groupsResponse.value || []) {
        try {
          // Try to get group calendar
          const calendarResponse = await this.graphClient
            .api(`/groups/${group.id}/calendar`)
            .select('id,name')
            .get();

          calendars.push({
            id: calendarResponse.id,
            title: `${calendarResponse.name} (Group)`,
            description: `Group calendar for ${group.displayName}`,
            type: CalendarSourceType.Exchange,
            url: 'https://outlook.office365.com/calendar/view/month',
            siteTitle: 'Exchange Online',
            siteUrl: 'https://outlook.office365.com',
            color: ColorUtils.generateColorFromString(group.displayName),
            isEnabled: true,
            canEdit: false, // Group calendars are typically read-only for members
            canShare: false
          });
        } catch {
          // Group might not have a calendar or user might not have access
          console.debug(`No calendar access for group ${group.displayName}`);
        }
      }
    } catch (error) {
      console.error('Error getting group calendars:', error);
    }

    return calendars;
  }

  /**
   * Get events from Exchange calendar
   */
  public async getEventsFromCalendar(source: ICalendarSource, maxEvents: number = 100): Promise<ICalendarEvent[]> {
    const events: ICalendarEvent[] = [];

    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        throw new Error('Graph client not available');
      }

      const now = new Date();
      const endDate = new Date();
      endDate.setMonth(endDate.getMonth() + 6); // Next 6 months

      let apiPath: string;
      if (source.title.includes('(Group)')) {
        // Extract group ID from source ID and use group calendar endpoint
        apiPath = `/groups/${source.id}/calendar/events`;
      } else {
        // Personal or shared calendar
        apiPath = `/me/calendars/${source.id}/events`;
      }

      const response = await this.graphClient
        .api(apiPath)
        .select(AppConstants.GRAPH_EVENT_FIELDS)
        .filter(`start/dateTime ge '${now.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`)
        .orderby('start/dateTime')
        .top(Math.min(maxEvents, AppConstants.API_LIMITS.MAX_EVENTS_PER_REQUEST))
        .get();

      for (const item of response.value || []) {
        events.push(this.mapGraphEventToCalendarEvent(item, source));
      }

    } catch (error) {
      console.error(`Error fetching Exchange events from ${source.title}:`, error);
      throw error;
    }

    return events;
  }

  /**
   * Map Microsoft Graph event to calendar event
   */
  private mapGraphEventToCalendarEvent(graphEvent: GraphEvent, source: ICalendarSource): ICalendarEvent {
    const startDate = new Date(graphEvent.start.dateTime + (graphEvent.start.timeZone ? '' : 'Z'));
    const endDate = new Date(graphEvent.end.dateTime + (graphEvent.end.timeZone ? '' : 'Z'));

    return {
      id: `ex_${source.id}_${graphEvent.id}`,
      title: graphEvent.subject || 'Untitled Event',
      description: this.extractTextFromHtml(graphEvent.body?.content || ''),
      start: startDate,
      end: endDate,
      location: graphEvent.location?.displayName || '',
      category: graphEvent.categories?.join(', ') || '',
      isAllDay: graphEvent.isAllDay || false,
      isRecurring: !!graphEvent.recurrence,
      calendarId: source.id,
      calendarTitle: source.title,
      calendarType: source.type,
      organizer: graphEvent.organizer?.emailAddress?.name || 'Unknown',
      created: new Date(graphEvent.createdDateTime),
      modified: new Date(graphEvent.lastModifiedDateTime),
      webUrl: graphEvent.webLink || 'https://outlook.office365.com/calendar',
      color: source.color || ColorUtils.generateColorFromString(source.title),
      importance: graphEvent.importance,
      sensitivity: graphEvent.sensitivity,
      showAs: graphEvent.showAs,
      attendees: this.mapAttendees(graphEvent.attendees || [])
    };
  }

  /**
   * Extract plain text from HTML content
   */
  private extractTextFromHtml(html: string): string {
    if (!html) return '';
    
    // Remove HTML tags and decode entities
    return html
      .replace(/<[^>]*>/g, '') // Remove HTML tags
      .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
      .replace(/&amp;/g, '&') // Replace encoded ampersands
      .replace(/&lt;/g, '<') // Replace encoded less than
      .replace(/&gt;/g, '>') // Replace encoded greater than
      .replace(/&quot;/g, '"') // Replace encoded quotes
      .trim();
  }

  /**
   * Map Graph attendees to calendar event attendees
   */
  private mapAttendees(graphAttendees: GraphEvent['attendees']): Array<{
    name: string;
    email: string;
    response: 'accepted' | 'declined' | 'tentative' | 'none';
    type: 'required' | 'optional' | 'resource';
  }> {
    if (!graphAttendees) return [];
    
    return graphAttendees.map(attendee => ({
      name: attendee.emailAddress?.name || 'Unknown',
      email: attendee.emailAddress?.address || '',
      response: (attendee.status?.response || 'none') as 'accepted' | 'declined' | 'tentative' | 'none',
      type: (attendee.type || 'required') as 'required' | 'optional' | 'resource'
    }));
  }

  /**
   * Search events across Exchange calendars
   */
  public async searchEvents(sources: ICalendarSource[], query: string, maxResults: number = 50): Promise<ICalendarEvent[]> {
    const allEvents: ICalendarEvent[] = [];

    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return allEvents;
      }

      for (const source of sources.filter(s => s.isEnabled && s.type === CalendarSourceType.Exchange)) {
        try {
          let apiPath: string;
          if (source.title.includes('(Group)')) {
            apiPath = `/groups/${source.id}/calendar/events`;
          } else {
            apiPath = `/me/calendars/${source.id}/events`;
          }

          const response = await this.graphClient
            .api(apiPath)
            .select(AppConstants.GRAPH_EVENT_FIELDS)
            .search(`"${query}"`)
            .top(maxResults)
            .get();

          for (const item of response.value || []) {
            allEvents.push(this.mapGraphEventToCalendarEvent(item, source));
          }
        } catch (error) {
          console.error(`Error searching events in ${source.title}:`, error);
        }
      }
    } catch (error) {
      console.error('Error in Exchange search:', error);
    }

    return allEvents
      .sort((a, b) => a.start.getTime() - b.start.getTime())
      .slice(0, maxResults);
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
    const allEvents: ICalendarEvent[] = [];

    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return allEvents;
      }

      for (const source of sources.filter(s => s.isEnabled && s.type === CalendarSourceType.Exchange)) {
        try {
          let apiPath: string;
          if (source.title.includes('(Group)')) {
            apiPath = `/groups/${source.id}/calendar/events`;
          } else {
            apiPath = `/me/calendars/${source.id}/events`;
          }

          const response = await this.graphClient
            .api(apiPath)
            .select(AppConstants.GRAPH_EVENT_FIELDS)
            .filter(`start/dateTime ge '${startDate.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`)
            .orderby('start/dateTime')
            .top(maxEvents)
            .get();

          for (const item of response.value || []) {
            allEvents.push(this.mapGraphEventToCalendarEvent(item, source));
          }
        } catch (error) {
          console.error(`Error getting events from ${source.title}:`, error);
        }
      }
    } catch (error) {
      console.error('Error getting Exchange events for date range:', error);
    }

    return allEvents
      .sort((a, b) => a.start.getTime() - b.start.getTime())
      .slice(0, maxEvents);
  }

  /**
   * Get calendar details
   */
  public async getCalendarDetails(calendarId: string): Promise<GraphCalendar | undefined> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return undefined;
      }

      const response = await this.graphClient
        .api(`/me/calendars/${calendarId}`)
        .select(AppConstants.GRAPH_CALENDAR_FIELDS)
        .get();

      return response;
    } catch (error) {
      console.error(`Error getting calendar details for ${calendarId}:`, error);
      return undefined;
    }
  }

  /**
   * Check if user has access to Graph API
   */
  public async checkGraphAccess(): Promise<boolean> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return false;
      }

      // Try to make a simple Graph call
      await this.graphClient
        .api('/me')
        .select('id')
        .get();

      return true;
    } catch (error) {
      console.warn('Graph API access check failed:', error);
      return false;
    }
  }

  /**
   * Get user's timezone
   */
  public async getUserTimeZone(): Promise<string> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return 'UTC';
      }

      const response = await this.graphClient
        .api('/me/mailboxSettings')
        .select('timeZone')
        .get();

      return response.timeZone || 'UTC';
    } catch (error) {
      console.warn('Could not get user timezone:', error);
      return 'UTC';
    }
  }

  /**
   * Get user's working hours
   */
  public async getUserWorkingHours(): Promise<Record<string, unknown> | undefined> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return undefined;
      }

      const response = await this.graphClient
        .api('/me/mailboxSettings')
        .select('workingHours')
        .get();

      return response.workingHours;
    } catch (error) {
      console.warn('Could not get user working hours:', error);
      return undefined;
    }
  }

  /**
   * Get free/busy information for a time period
   */
  public async getFreeBusyInfo(calendarIds: string[], startDate: Date, endDate: Date): Promise<Record<string, unknown>[] | undefined> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return undefined;
      }

      const requestBody = {
        schedules: calendarIds,
        startTime: {
          dateTime: startDate.toISOString(),
          timeZone: 'UTC'
        },
        endTime: {
          dateTime: endDate.toISOString(),
          timeZone: 'UTC'
        },
        availabilityViewInterval: 15 // 15-minute intervals
      };

      const response = await this.graphClient
        .api('/me/calendar/getSchedule')
        .post(requestBody);

      return response.value;
    } catch (error) {
      console.warn('Could not get free/busy information:', error);
      return undefined;
    }
  }

  /**
   * Create a new event in Exchange calendar
   */
  public async createEvent(calendarId: string, eventData: Record<string, unknown>): Promise<ICalendarEvent | undefined> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        throw new Error('Graph client not available');
      }

      const response = await this.graphClient
        .api(`/me/calendars/${calendarId}/events`)
        .post(eventData);

      // Find the calendar source to map the event correctly
      const calendarSources = await this.getExchangeCalendars();
      const source = calendarSources.find(s => s.id === calendarId);
      
      if (source) {
        return this.mapGraphEventToCalendarEvent(response, source);
      }

      return undefined;
    } catch (error) {
      console.error('Error creating event:', error);
      throw error;
    }
  }

  /**
   * Update an existing event
   */
  public async updateEvent(calendarId: string, eventId: string, eventData: Record<string, unknown>): Promise<ICalendarEvent | undefined> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        throw new Error('Graph client not available');
      }

      const response = await this.graphClient
        .api(`/me/calendars/${calendarId}/events/${eventId}`)
        .patch(eventData);

      // Find the calendar source to map the event correctly
      const calendarSources = await this.getExchangeCalendars();
      const source = calendarSources.find(s => s.id === calendarId);
      
      if (source) {
        return this.mapGraphEventToCalendarEvent(response, source);
      }

      return undefined;
    } catch (error) {
      console.error('Error updating event:', error);
      throw error;
    }
  }

  /**
   * Delete an event
   */
  public async deleteEvent(calendarId: string, eventId: string): Promise<boolean> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        throw new Error('Graph client not available');
      }

      await this.graphClient
        .api(`/me/calendars/${calendarId}/events/${eventId}`)
        .delete();

      return true;
    } catch (error) {
      console.error('Error deleting event:', error);
      throw error;
    }
  }

  /**
   * Get calendar permissions for current user
   */
  public async getCalendarPermissions(calendarId: string): Promise<{
    canEdit: boolean;
    canShare: boolean;
    canViewPrivateItems: boolean;
  }> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return {
          canEdit: false,
          canShare: false,
          canViewPrivateItems: false
        };
      }

      const response = await this.graphClient
        .api(`/me/calendars/${calendarId}`)
        .select('canEdit,canShare,canViewPrivateItems')
        .get();

      return {
        canEdit: response.canEdit || false,
        canShare: response.canShare || false,
        canViewPrivateItems: response.canViewPrivateItems || false
      };
    } catch (error) {
      console.warn(`Could not get permissions for calendar ${calendarId}:`, error);
      return {
        canEdit: false,
        canShare: false,
        canViewPrivateItems: false
      };
    }
  }

  /**
   * Get calendar statistics
   */
  public async getCalendarStatistics(calendarId: string): Promise<{ totalEvents: number; upcomingEvents: number }> {
    try {
      await this.initializeGraphClient();
      
      if (!this.graphClient) {
        return { totalEvents: 0, upcomingEvents: 0 };
      }

      const now = new Date();
      const oneMonthFromNow = new Date();
      oneMonthFromNow.setMonth(oneMonthFromNow.getMonth() + 1);

      // Get upcoming events count
      const upcomingResponse = await this.graphClient
        .api(`/me/calendars/${calendarId}/events`)
        .filter(`start/dateTime ge '${now.toISOString()}' and start/dateTime le '${oneMonthFromNow.toISOString()}'`)
        .select('id')
        .get();

      // Get total events count (approximate)
      const totalResponse = await this.graphClient
        .api(`/me/calendars/${calendarId}/events`)
        .select('id')
        .top(1000)
        .get();

      return {
        totalEvents: totalResponse.value?.length || 0,
        upcomingEvents: upcomingResponse.value?.length || 0
      };
    } catch (error) {
      console.warn(`Could not get statistics for calendar ${calendarId}:`, error);
      return { totalEvents: 0, upcomingEvents: 0 };
    }
  }
}