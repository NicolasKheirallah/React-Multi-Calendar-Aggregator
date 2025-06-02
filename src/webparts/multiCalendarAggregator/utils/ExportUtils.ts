import { ICalendarEvent } from '../models/ICalendarModels';
import { DateUtils } from './DateUtils';
import { AppConstants } from '../constants/AppConstants';

export class ExportUtils {
  /**
   * Export events to ICS format
   */
  public static exportToICS(events: ICalendarEvent[], filename?: string): void {
    const icsContent = this.generateICSContent(events);
    this.downloadFile(icsContent, filename || 'calendar-events.ics', 'text/calendar');
  }

  /**
   * Export events to CSV format
   */
  public static exportToCSV(events: ICalendarEvent[], filename?: string): void {
    const csvContent = this.generateCSVContent(events);
    this.downloadFile(csvContent, filename || 'calendar-events.csv', 'text/csv');
  }

  /**
   * Export events to JSON format
   */
  public static exportToJSON(events: ICalendarEvent[], filename?: string): void {
    const jsonContent = this.generateJSONContent(events);
    this.downloadFile(jsonContent, filename || 'calendar-events.json', 'application/json');
  }

  /**
   * Export events to Excel format (simplified XLSX)
   */
  public static exportToExcel(events: ICalendarEvent[], filename?: string): void {
    // For a full XLSX implementation, you would use a library like xlsx
    // This is a simplified CSV-based approach that Excel can open
    const csvContent = this.generateExcelCompatibleCSV(events);
    this.downloadFile(csvContent, filename || 'calendar-events.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  /**
   * Generate ICS (iCalendar) content
   */
  private static generateICSContent(events: ICalendarEvent[]): string {
    const icsLines: string[] = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Multi-Calendar Aggregator//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH'
    ];

    events.forEach(event => {
      icsLines.push(...this.generateICSEvent(event));
    });

    icsLines.push('END:VCALENDAR');

    return icsLines.join('\r\n');
  }

  /**
   * Generate individual ICS event
   */
  private static generateICSEvent(event: ICalendarEvent): string[] {
    const lines: string[] = [];
    const now = new Date();
    const uid = `${event.id}@multi-calendar-aggregator`;

    lines.push('BEGIN:VEVENT');
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${this.formatICSDateTime(now)}`);
    lines.push(`DTSTART:${this.formatICSDateTime(event.start)}`);
    lines.push(`DTEND:${this.formatICSDateTime(event.end)}`);
    lines.push(`SUMMARY:${this.escapeICSText(event.title)}`);

    if (event.description) {
      lines.push(`DESCRIPTION:${this.escapeICSText(event.description)}`);
    }

    if (event.location) {
      lines.push(`LOCATION:${this.escapeICSText(event.location)}`);
    }

    if (event.organizer) {
      lines.push(`ORGANIZER:CN=${this.escapeICSText(event.organizer)}`);
    }

    if (event.category) {
      lines.push(`CATEGORIES:${this.escapeICSText(event.category)}`);
    }

    if (event.webUrl) {
      lines.push(`URL:${event.webUrl}`);
    }

    lines.push(`CREATED:${this.formatICSDateTime(event.created)}`);
    lines.push(`LAST-MODIFIED:${this.formatICSDateTime(event.modified)}`);

    if (event.isAllDay) {
      lines.push('X-MICROSOFT-CDO-ALLDAYEVENT:TRUE');
    }

    if (event.importance) {
      const priority = event.importance === 'high' ? '1' : event.importance === 'low' ? '9' : '5';
      lines.push(`PRIORITY:${priority}`);
    }

    lines.push(`X-CALENDAR-SOURCE:${event.calendarTitle}`);
    lines.push(`X-CALENDAR-TYPE:${event.calendarType}`);

    lines.push('END:VEVENT');

    return lines;
  }

  /**
   * Format date for ICS
   */
  private static formatICSDateTime(date: Date): string {
    return date.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  }

  /**
   * Escape text for ICS format
   */
  private static escapeICSText(text: string): string {
    return text
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '');
  }

  /**
   * Generate CSV content
   */
  private static generateCSVContent(events: ICalendarEvent[]): string {
    const headers = [
      'Title',
      'Description',
      'Start Date',
      'Start Time',
      'End Date',
      'End Time',
      'Location',
      'Category',
      'All Day',
      'Recurring',
      'Calendar',
      'Calendar Type',
      'Organizer',
      'Created',
      'Modified',
      'Web URL'
    ];

    const csvLines: string[] = [headers.join(',')];

    events.forEach(event => {
      const row = [
        this.escapeCSVField(event.title),
        this.escapeCSVField(event.description || ''),
        DateUtils.formatDate(event.start),
        event.isAllDay ? 'All Day' : DateUtils.formatTime(event.start),
        DateUtils.formatDate(event.end),
        event.isAllDay ? 'All Day' : DateUtils.formatTime(event.end),
        this.escapeCSVField(event.location || ''),
        this.escapeCSVField(event.category || ''),
        event.isAllDay ? 'Yes' : 'No',
        event.isRecurring ? 'Yes' : 'No',
        this.escapeCSVField(event.calendarTitle),
        event.calendarType,
        this.escapeCSVField(event.organizer),
        DateUtils.formatDateTime(event.created),
        DateUtils.formatDateTime(event.modified),
        event.webUrl || ''
      ];

      csvLines.push(row.join(','));
    });

    return csvLines.join('\n');
  }

  /**
   * Escape CSV field
   */
  private static escapeCSVField(field: string): string {
    if (!field) return '""';
    
    // If field contains comma, quote, or newline, wrap in quotes and escape quotes
    if (field.includes(',') || field.includes('"') || field.includes('\n')) {
      return `"${field.replace(/"/g, '""')}"`;
    }
    
    return field;
  }

  /**
   * Generate Excel-compatible CSV content
   */
  private static generateExcelCompatibleCSV(events: ICalendarEvent[]): string {
    const headers = [
      'Subject',
      'Start Date',
      'Start Time',
      'End Date', 
      'End Time',
      'All day event',
      'Description',
      'Location',
      'Categories',
      'Show time as',
      'Organizer',
      'Required Attendees',
      'Optional Attendees',
      'Meeting Resources',
      'Billing Information',
      'Mileage',
      'Priority',
      'Private',
      'Sensitivity',
      'Recurrence Pattern',
      'Calendar',
      'Calendar Type'
    ];

    const csvLines: string[] = [headers.join(',')];

    events.forEach(event => {
      const row = [
        this.escapeCSVField(event.title),
        DateUtils.formatDate(event.start, 'MM/DD/YYYY'),
        event.isAllDay ? '' : DateUtils.formatTime(event.start, 'HH:mm:ss'),
        DateUtils.formatDate(event.end, 'MM/DD/YYYY'),
        event.isAllDay ? '' : DateUtils.formatTime(event.end, 'HH:mm:ss'),
        event.isAllDay ? 'True' : 'False',
        this.escapeCSVField(event.description || ''),
        this.escapeCSVField(event.location || ''),
        this.escapeCSVField(event.category || ''),
        event.showAs || 'Busy',
        this.escapeCSVField(event.organizer),
        '', // Required Attendees - would need to parse from attendees
        '', // Optional Attendees
        '', // Meeting Resources
        '', // Billing Information
        '', // Mileage
        event.importance || 'Normal',
        'False', // Private
        event.sensitivity || 'Normal',
        event.isRecurring ? 'Yes' : '',
        this.escapeCSVField(event.calendarTitle),
        event.calendarType
      ];

      csvLines.push(row.join(','));
    });

    return '\ufeff' + csvLines.join('\r\n'); // Add BOM for Excel UTF-8 support
  }

  /**
   * Generate JSON content
   */
  private static generateJSONContent(events: ICalendarEvent[]): string {
    const exportData = {
      metadata: {
        exportDate: new Date().toISOString(),
        totalEvents: events.length,
        exportedBy: AppConstants.APP_NAME,
        version: AppConstants.APP_VERSION
      },
      events: events.map(event => ({
        id: event.id,
        title: event.title,
        description: event.description,
        startDateTime: event.start.toISOString(),
        endDateTime: event.end.toISOString(),
        location: event.location,
        category: event.category,
        isAllDay: event.isAllDay,
        isRecurring: event.isRecurring,
        calendar: {
          id: event.calendarId,
          title: event.calendarTitle,
          type: event.calendarType
        },
        organizer: event.organizer,
        createdDateTime: event.created.toISOString(),
        modifiedDateTime: event.modified.toISOString(),
        webUrl: event.webUrl,
        color: event.color,
        importance: event.importance,
        sensitivity: event.sensitivity,
        showAs: event.showAs,
        attendees: event.attendees || []
      }))
    };

    return JSON.stringify(exportData, null, 2);
  }

  /**
   * Download file to user's device
   */
  private static downloadFile(content: string, filename: string, mimeType: string): void {
    try {
      const blob = new Blob([content], { type: mimeType });
      const url = window.URL.createObjectURL(blob);
      
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      link.style.display = 'none';
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      // Clean up the URL object
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error downloading file:', error);
      
      // Fallback: open in new window
      const dataUrl = `data:${mimeType};charset=utf-8,${encodeURIComponent(content)}`;
      window.open(dataUrl, '_blank');
    }
  }

  /**
   * Generate summary report
   */
  public static generateSummaryReport(events: ICalendarEvent[]): string {
    const now = new Date();
    const report: string[] = [];

    report.push('# Calendar Events Summary Report');
    report.push('');
    report.push(`**Generated:** ${DateUtils.formatDateTime(now)}`);
    report.push(`**Total Events:** ${events.length}`);
    report.push('');

    // Events by calendar
    const eventsByCalendar = this.groupEventsByCalendar(events);
    report.push('## Events by Calendar');
    report.push('');
    Object.entries(eventsByCalendar).forEach(([calendar, calendarEvents]) => {
      report.push(`- **${calendar}:** ${calendarEvents.length} events`);
    });
    report.push('');

    // Events by date range
    const eventsByMonth = this.groupEventsByMonth(events);
    report.push('## Events by Month');
    report.push('');
    Object.entries(eventsByMonth).forEach(([month, monthEvents]) => {
      report.push(`- **${month}:** ${monthEvents.length} events`);
    });
    report.push('');

    // Upcoming events (next 7 days)
    const upcomingEvents = events.filter(event => {
      const daysDiff = (event.start.getTime() - now.getTime()) / (1000 * 60 * 60 * 24);
      return daysDiff >= 0 && daysDiff <= 7;
    });

    report.push('## Upcoming Events (Next 7 Days)');
    report.push('');
    if (upcomingEvents.length === 0) {
      report.push('No upcoming events in the next 7 days.');
    } else {
      upcomingEvents.slice(0, 10).forEach(event => {
        const dateStr = DateUtils.formatDateTime(event.start);
        report.push(`- **${event.title}** - ${dateStr}`);
        if (event.location) {
          report.push(`  - Location: ${event.location}`);
        }
      });
      
      if (upcomingEvents.length > 10) {
        report.push(`  - ... and ${upcomingEvents.length - 10} more events`);
      }
    }
    report.push('');

    // Statistics
    const allDayEvents = events.filter(e => e.isAllDay).length;
    const recurringEvents = events.filter(e => e.isRecurring).length;
    const eventsWithLocation = events.filter(e => e.location && e.location.trim()).length;

    report.push('## Statistics');
    report.push('');
    report.push(`- **All-day events:** ${allDayEvents} (${Math.round(allDayEvents / events.length * 100)}%)`);
    report.push(`- **Recurring events:** ${recurringEvents} (${Math.round(recurringEvents / events.length * 100)}%)`);
    report.push(`- **Events with location:** ${eventsWithLocation} (${Math.round(eventsWithLocation / events.length * 100)}%)`);

    return report.join('\n');
  }

  /**
   * Group events by calendar
   */
  private static groupEventsByCalendar(events: ICalendarEvent[]): { [calendar: string]: ICalendarEvent[] } {
    return events.reduce((groups, event) => {
      const calendar = event.calendarTitle;
      if (!groups[calendar]) {
        groups[calendar] = [];
      }
      groups[calendar].push(event);
      return groups;
    }, {} as { [calendar: string]: ICalendarEvent[] });
  }

  /**
   * Group events by month
   */
  private static groupEventsByMonth(events: ICalendarEvent[]): { [month: string]: ICalendarEvent[] } {
    return events.reduce((groups, event) => {
      const month = DateUtils.formatDate(event.start, 'MMMM YYYY');
      if (!groups[month]) {
        groups[month] = [];
      }
      groups[month].push(event);
      return groups;
    }, {} as { [month: string]: ICalendarEvent[] });
  }

  /**
   * Export filtered events based on criteria
   */
  public static exportFilteredEvents(
    events: ICalendarEvent[],
    filters: {
      startDate?: Date;
      endDate?: Date;
      calendars?: string[];
      categories?: string[];
      searchQuery?: string;
    },
    format: 'ics' | 'csv' | 'json' | 'excel' = 'ics',
    filename?: string
  ): void {
    let filteredEvents = [...events];

    // Apply date filters
    if (filters.startDate) {
      filteredEvents = filteredEvents.filter(event => event.start >= filters.startDate!);
    }
    if (filters.endDate) {
      filteredEvents = filteredEvents.filter(event => event.start <= filters.endDate!);
    }

    // Apply calendar filters
    if (filters.calendars && filters.calendars.length > 0) {
      filteredEvents = filteredEvents.filter(event => 
        filters.calendars!.includes(event.calendarId)
      );
    }

    // Apply category filters
    if (filters.categories && filters.categories.length > 0) {
      filteredEvents = filteredEvents.filter(event => 
        event.category && filters.categories!.includes(event.category)
      );
    }

    // Apply search query
    if (filters.searchQuery) {
      const query = filters.searchQuery.toLowerCase();
      filteredEvents = filteredEvents.filter(event =>
        event.title.toLowerCase().includes(query) ||
        (event.description && event.description.toLowerCase().includes(query)) ||
        (event.location && event.location.toLowerCase().includes(query))
      );
    }

    // Export based on format
    switch (format) {
      case 'csv':
        this.exportToCSV(filteredEvents, filename);
        break;
      case 'json':
        this.exportToJSON(filteredEvents, filename);
        break;
      case 'excel':
        this.exportToExcel(filteredEvents, filename);
        break;
      case 'ics':
      default:
        this.exportToICS(filteredEvents, filename);
        break;
    }
  }

  /**
   * Create shareable event URL
   */
  public static createShareableEventURL(event: ICalendarEvent): string {
    const params = new URLSearchParams({
      action: 'TEMPLATE',
      text: event.title,
      dates: `${this.formatGoogleDate(event.start)}/${this.formatGoogleDate(event.end)}`,
      details: event.description || '',
      location: event.location || '',
      sprop: 'website:multi-calendar-aggregator'
    });

    return `https://calendar.google.com/calendar/render?${params.toString()}`;
  }

  /**
   * Format date for Google Calendar
   */
  private static formatGoogleDate(date: Date): string {
    return date.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  }

  /**
   * Generate calendar subscription URL (for ICS feed)
   */
  public static generateSubscriptionURL(calendarIds: string[], baseUrl: string): string {
    const params = new URLSearchParams({
      calendars: calendarIds.join(','),
      format: 'ics'
    });

    return `${baseUrl}/api/calendar-feed?${params.toString()}`;
  }

  /**
   * Export calendar statistics
   */
  public static exportStatistics(events: ICalendarEvent[], filename?: string): void {
    const stats = this.generateStatistics(events);
    const csvContent = this.statisticsToCSV(stats);
    this.downloadFile(csvContent, filename || 'calendar-statistics.csv', 'text/csv');
  }

  /**
   * Generate detailed statistics
   */
  private static generateStatistics(events: ICalendarEvent[]): any {
    const now = new Date();
    
    return {
      overview: {
        totalEvents: events.length,
        upcomingEvents: events.filter(e => e.start >= now).length,
        pastEvents: events.filter(e => e.start < now).length,
        allDayEvents: events.filter(e => e.isAllDay).length,
        recurringEvents: events.filter(e => e.isRecurring).length
      },
      byCalendar: this.groupEventsByCalendar(events),
      byMonth: this.groupEventsByMonth(events),
      byCategory: events.reduce((acc, event) => {
        const category = event.category || 'Uncategorized';
        acc[category] = (acc[category] || 0) + 1;
        return acc;
      }, {} as { [key: string]: number }),
      byDayOfWeek: events.reduce((acc, event) => {
        const dayName = DateUtils.formatDate(event.start, 'dddd');
        acc[dayName] = (acc[dayName] || 0) + 1;
        return acc;
      }, {} as { [key: string]: number })
    };
  }

  /**
   * Convert statistics to CSV
   */
  private static statisticsToCSV(stats: any): string {
    const lines: string[] = [];
    
    lines.push('Calendar Statistics Report');
    lines.push('');
    
    // Overview
    lines.push('Overview');
    lines.push('Metric,Count');
    Object.entries(stats.overview).forEach(([key, value]) => {
      lines.push(`${key},${value}`);
    });
    lines.push('');
    
    // By Calendar
    lines.push('Events by Calendar');
    lines.push('Calendar,Count');
    Object.entries(stats.byCalendar).forEach(([calendar, events]: [string, any]) => {
      lines.push(`${this.escapeCSVField(calendar)},${events.length}`);
    });
    lines.push('');
    
    // By Category
    lines.push('Events by Category');
    lines.push('Category,Count');
    Object.entries(stats.byCategory).forEach(([category, count]) => {
      lines.push(`${this.escapeCSVField(category)},${count}`);
    });
    
    return lines.join('\n');
  }
}