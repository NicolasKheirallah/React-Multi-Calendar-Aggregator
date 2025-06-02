import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICalendarEvent, ICalendarSource, CalendarSourceType } from '../models/ICalendarModels';
import { AppConstants } from '../constants/AppConstants';
import { DateUtils } from '../utils/DateUtils';
import { ColorUtils } from '../utils/ColorUtils';

interface SharePointListItem {
  [key: string]: unknown;
  Id: number;
  Title: string;
  Description?: string;
  EventDate: string;
  EndDate?: string;
  Location?: string;
  Category?: string;
  fAllDayEvent?: boolean;
  fRecurrence?: boolean;
  RecurrenceData?: string;
  Created: string;
  Modified: string;
  Author?: {
    Title: string;
  };
  Editor?: {
    Title: string;
  };
}

interface SharePointList {
  Id: string;
  Title: string;
  Description?: string;
  DefaultViewUrl: string;
  ItemCount?: number;
  LastItemModifiedDate?: string;
  Created: string;
  RootFolder?: {
    ServerRelativeUrl: string;
  };
}

interface SharePointSite {
  Id: string;
  Title: string;
  Url: string;
  Description?: string;
  Created: string;
}

export class SharePointCalendarService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Get all SharePoint calendar lists from current site and subsites
   */
  public async getSharePointCalendars(): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      // Get calendars from current site
      const currentSiteCalendars = await this.getCalendarsFromSite(
        this.context.pageContext.web.absoluteUrl,
        this.context.pageContext.web.title || 'Current Site'
      );
      calendars.push(...currentSiteCalendars);

      // Get calendars from subsites
      const subsites = await this.getSubsites();
      for (const subsite of subsites) {
        try {
          const subsiteCalendars = await this.getCalendarsFromSite(subsite.Url, subsite.Title);
          calendars.push(...subsiteCalendars);
        } catch (error) {
          console.warn(`Could not access calendars from subsite ${subsite.Title}:`, error);
        }
      }

      // Get calendars from parent site if this is a subsite
      if (this.context.pageContext.web.absoluteUrl !== this.context.pageContext.site.absoluteUrl) {
        try {
          const parentSiteCalendars = await this.getCalendarsFromSite(
            this.context.pageContext.site.absoluteUrl,
            'Parent Site'
          );
          calendars.push(...parentSiteCalendars);
        } catch (error) {
          console.warn('Could not access parent site calendars:', error);
        }
      }

    } catch (error) {
      console.error('Error getting SharePoint calendars:', error);
      throw new Error(AppConstants.ERROR_MESSAGES.SHAREPOINT_API_ERROR);
    }

    return calendars;
  }

  /**
   * Get calendar lists from a specific SharePoint site
   */
  private async getCalendarsFromSite(siteUrl: string, siteName: string): Promise<ICalendarSource[]> {
    const calendars: ICalendarSource[] = [];

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${siteUrl}${AppConstants.SHAREPOINT_API.LISTS_ENDPOINT}?` +
        `$filter=${AppConstants.SHAREPOINT_API.CALENDAR_FILTER}&` +
        `$select=Id,Title,DefaultViewUrl,Description,ItemCount,LastItemModifiedDate,Created,RootFolder/ServerRelativeUrl&` +
        `$expand=RootFolder`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const lists = data.d?.results || data.value || [];

        for (const list of lists as SharePointList[]) {
          calendars.push({
            id: list.Id,
            title: list.Title,
            description: list.Description || `SharePoint calendar from ${siteName}`,
            type: CalendarSourceType.SharePoint,
            url: `${siteUrl}${list.DefaultViewUrl}`,
            siteTitle: siteName,
            siteUrl: siteUrl,
            itemCount: list.ItemCount || 0,
            lastModified: list.LastItemModifiedDate,
            color: ColorUtils.generateColorFromString(list.Title),
            isEnabled: true,
            canEdit: true, // Will be determined by actual permissions
            canShare: true
          });
        }
      } else {
        console.warn(`Failed to get calendars from ${siteUrl}. Status: ${response.status}`);
      }
    } catch (error) {
      console.warn(`Could not access calendars from ${siteUrl}:`, error);
    }

    return calendars;
  }

  /**
   * Get subsites from current site
   */
  private async getSubsites(): Promise<SharePointSite[]> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}${AppConstants.SHAREPOINT_API.SITES_ENDPOINT}?` +
        `$select=Id,Title,Url,Description,Created`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        return data.d?.results || data.value || [];
      }
    } catch (error) {
      console.warn('Could not get subsites:', error);
    }

    return [];
  }

  /**
   * Get events from a SharePoint calendar list
   */
  public async getEventsFromCalendar(source: ICalendarSource, maxEvents: number = 100): Promise<ICalendarEvent[]> {
    const events: ICalendarEvent[] = [];
    
    try {
      const now = new Date();
      const endDate = DateUtils.addTime(now, 6, 'months');
      
      // Build the query
      const filter = `${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE} ge datetime'${now.toISOString()}' and ` +
                    `${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE} le datetime'${endDate.toISOString()}'`;
      
      const select = [
        AppConstants.SHAREPOINT_FIELDS.ID,
        AppConstants.SHAREPOINT_FIELDS.TITLE,
        AppConstants.SHAREPOINT_FIELDS.DESCRIPTION,
        AppConstants.SHAREPOINT_FIELDS.EVENT_DATE,
        AppConstants.SHAREPOINT_FIELDS.END_DATE,
        AppConstants.SHAREPOINT_FIELDS.LOCATION,
        AppConstants.SHAREPOINT_FIELDS.CATEGORY,
        AppConstants.SHAREPOINT_FIELDS.ALL_DAY_EVENT,
        AppConstants.SHAREPOINT_FIELDS.RECURRENCE,
        AppConstants.SHAREPOINT_FIELDS.RECURRENCE_DATA,
        AppConstants.SHAREPOINT_FIELDS.CREATED,
        AppConstants.SHAREPOINT_FIELDS.MODIFIED,
        'Author/Title',
        'Editor/Title'
      ].join(',');

      const apiUrl = `${source.siteUrl}${AppConstants.SHAREPOINT_API.LIST_ITEMS_ENDPOINT.replace('{listId}', source.id)}?` +
        `$select=${select}&` +
        `$expand=Author,Editor&` +
        `$filter=${filter}&` +
        `$orderby=${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE}&` +
        `$top=${Math.min(maxEvents, AppConstants.API_LIMITS.MAX_EVENTS_PER_REQUEST)}`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.d?.results || data.value || [];

        for (const item of items as SharePointListItem[]) {
          const event = this.mapSharePointItemToEvent(item, source);
          events.push(event);

          // Handle recurring events
          if (item[AppConstants.SHAREPOINT_FIELDS.RECURRENCE] && item[AppConstants.SHAREPOINT_FIELDS.RECURRENCE_DATA]) {
            const recurringEvents = this.expandRecurringEvent(event, item[AppConstants.SHAREPOINT_FIELDS.RECURRENCE_DATA] as string, endDate);
            events.push(...recurringEvents);
          }
        }
      } else {
        throw new Error(`Failed to fetch events. Status: ${response.status}`);
      }

    } catch (error) {
      console.error(`Error fetching SharePoint events from ${source.title}:`, error);
      throw error;
    }

    return events.slice(0, maxEvents);
  }

  /**
   * Map SharePoint list item to calendar event
   */
  private mapSharePointItemToEvent(item: SharePointListItem, source: ICalendarSource): ICalendarEvent {
    const startDate = new Date(item[AppConstants.SHAREPOINT_FIELDS.EVENT_DATE] as string);
    const endDate = item[AppConstants.SHAREPOINT_FIELDS.END_DATE] 
      ? new Date(item[AppConstants.SHAREPOINT_FIELDS.END_DATE] as string)
      : startDate;

    return {
      id: `sp_${source.id}_${item[AppConstants.SHAREPOINT_FIELDS.ID]}`,
      title: (item[AppConstants.SHAREPOINT_FIELDS.TITLE] as string) || 'Untitled Event',
      description: (item[AppConstants.SHAREPOINT_FIELDS.DESCRIPTION] as string) || '',
      start: startDate,
      end: endDate,
      location: (item[AppConstants.SHAREPOINT_FIELDS.LOCATION] as string) || '',
      category: (item[AppConstants.SHAREPOINT_FIELDS.CATEGORY] as string) || '',
      isAllDay: (item[AppConstants.SHAREPOINT_FIELDS.ALL_DAY_EVENT] as boolean) || false,
      isRecurring: (item[AppConstants.SHAREPOINT_FIELDS.RECURRENCE] as boolean) || false,
      calendarId: source.id,
      calendarTitle: source.title,
      calendarType: source.type,
      organizer: item.Author?.Title || 'Unknown',
      created: new Date(item[AppConstants.SHAREPOINT_FIELDS.CREATED] as string),
      modified: new Date(item[AppConstants.SHAREPOINT_FIELDS.MODIFIED] as string),
      webUrl: `${source.siteUrl}/Lists/${source.title.replace(/\s+/g, '')}/DispForm.aspx?ID=${item[AppConstants.SHAREPOINT_FIELDS.ID]}`,
      color: source.color || ColorUtils.generateColorFromString(source.title)
    };
  }

  /**
   * Expand recurring events (simplified implementation)
   */
  private expandRecurringEvent(baseEvent: ICalendarEvent, recurrenceData: string, endDate: Date): ICalendarEvent[] {
    const recurringEvents: ICalendarEvent[] = [];
    
    try {
      // This is a simplified implementation for weekly recurrence
      // In production, you would parse the full SharePoint recurrence XML
      const eventDuration = baseEvent.end.getTime() - baseEvent.start.getTime();
      let currentDate = new Date(baseEvent.start);
      
      // Generate up to 50 recurring instances or until endDate
      for (let i = 1; i < 50 && currentDate < endDate; i++) {
        // Assume weekly recurrence for simplicity
        currentDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000); // Add 7 days properly
        
        if (currentDate > endDate) break;
        
        recurringEvents.push({
          ...baseEvent,
          id: `${baseEvent.id}_recur_${i}`,
          start: new Date(currentDate),
          end: new Date(currentDate.getTime() + eventDuration)
        });
      }
    } catch (error) {
      console.warn('Error expanding recurring event:', error);
    }

    return recurringEvents;
  }

  /**
   * Search events across SharePoint calendars
   */
  public async searchEvents(sources: ICalendarSource[], query: string, maxResults: number = 50): Promise<ICalendarEvent[]> {
    const allEvents: ICalendarEvent[] = [];
    const searchTerms = query.toLowerCase().split(' ').filter(term => term.length > 0);

    if (searchTerms.length === 0) {
      return allEvents;
    }

    for (const source of sources.filter(s => s.isEnabled && s.type === CalendarSourceType.SharePoint)) {
      try {
        // Build search filter for SharePoint
        const searchFilter = searchTerms.map(term => 
          `substringof('${term}',${AppConstants.SHAREPOINT_FIELDS.TITLE}) or ` +
          `substringof('${term}',${AppConstants.SHAREPOINT_FIELDS.DESCRIPTION}) or ` +
          `substringof('${term}',${AppConstants.SHAREPOINT_FIELDS.LOCATION})`
        ).join(' and ');

        const apiUrl = `${source.siteUrl}${AppConstants.SHAREPOINT_API.LIST_ITEMS_ENDPOINT.replace('{listId}', source.id)}?` +
          `$select=${AppConstants.SHAREPOINT_FIELDS.ID},${AppConstants.SHAREPOINT_FIELDS.TITLE},${AppConstants.SHAREPOINT_FIELDS.DESCRIPTION},${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE},${AppConstants.SHAREPOINT_FIELDS.END_DATE},${AppConstants.SHAREPOINT_FIELDS.LOCATION}&` +
          `$filter=${searchFilter}&` +
          `$orderby=${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE}&` +
          `$top=${maxResults}`;

        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          const items = data.d?.results || data.value || [];

          for (const item of items as SharePointListItem[]) {
            const event = this.mapSharePointItemToEvent(item, source);
            allEvents.push(event);
          }
        }
      } catch (error) {
        console.error(`Error searching events in ${source.title}:`, error);
      }
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

    for (const source of sources.filter(s => s.isEnabled && s.type === CalendarSourceType.SharePoint)) {
      try {
        const filter = `${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE} ge datetime'${startDate.toISOString()}' and ` +
                      `${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE} le datetime'${endDate.toISOString()}'`;

        const apiUrl = `${source.siteUrl}${AppConstants.SHAREPOINT_API.LIST_ITEMS_ENDPOINT.replace('{listId}', source.id)}?` +
          `$select=${AppConstants.SHAREPOINT_FIELDS.ID},${AppConstants.SHAREPOINT_FIELDS.TITLE},${AppConstants.SHAREPOINT_FIELDS.DESCRIPTION},${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE},${AppConstants.SHAREPOINT_FIELDS.END_DATE},${AppConstants.SHAREPOINT_FIELDS.LOCATION},${AppConstants.SHAREPOINT_FIELDS.CATEGORY}&` +
          `$filter=${filter}&` +
          `$orderby=${AppConstants.SHAREPOINT_FIELDS.EVENT_DATE}&` +
          `$top=${maxEvents}`;

        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          const items = data.d?.results || data.value || [];

          for (const item of items as SharePointListItem[]) {
            const event = this.mapSharePointItemToEvent(item, source);
            allEvents.push(event);
          }
        }
      } catch (error) {
        console.error(`Error getting events from ${source.title}:`, error);
      }
    }

    return allEvents
      .sort((a, b) => a.start.getTime() - b.start.getTime())
      .slice(0, maxEvents);
  }

  /**
   * Check if user has permissions to access a calendar
   */
  public async checkCalendarPermissions(source: ICalendarSource): Promise<{ canRead: boolean; canWrite: boolean }> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${source.siteUrl}/_api/web/lists(guid'${source.id}')/EffectiveBasePermissions`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const permissions = data.d?.EffectiveBasePermissions || data.EffectiveBasePermissions || {};
        
        // Check for specific permissions (simplified)
        const canRead = true; // If we can make the call, we can read
        const canWrite = permissions.High && permissions.Low; // Simplified check
        
        return { canRead, canWrite };
      }
    } catch (error) {
      console.warn(`Could not check permissions for ${source.title}:`, error);
    }

    return { canRead: false, canWrite: false };
  }

  /**
   * Get calendar statistics
   */
  public async getCalendarStatistics(source: ICalendarSource): Promise<{ totalItems: number; recentItems: number }> {
    try {
      const now = new Date();
      const oneWeekAgo = DateUtils.subtractTime(now, 7, 'days');

      // Get total count
      const totalResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${source.siteUrl}${AppConstants.SHAREPOINT_API.LIST_ITEMS_ENDPOINT.replace('{listId}', source.id)}?$select=Id&$top=1`,
        SPHttpClient.configurations.v1
      );

      // Get recent count
      const recentResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${source.siteUrl}${AppConstants.SHAREPOINT_API.LIST_ITEMS_ENDPOINT.replace('{listId}', source.id)}?` +
        `$select=Id&$filter=${AppConstants.SHAREPOINT_FIELDS.CREATED} ge datetime'${oneWeekAgo.toISOString()}'&$top=1000`,
        SPHttpClient.configurations.v1
      );

      let totalItems = 0;
      let recentItems = 0;

      if (totalResponse.ok) {
        const totalData = await totalResponse.json();
        totalItems = totalData.d?.__count || totalData['odata.count'] || 0;
      }

      if (recentResponse.ok) {
        const recentData = await recentResponse.json();
        recentItems = (recentData.d?.results || recentData.value || []).length;
      }

      return { totalItems, recentItems };
    } catch (error) {
      console.warn(`Could not get statistics for ${source.title}:`, error);
      return { totalItems: 0, recentItems: 0 };
    }
  }
}