import moment from 'moment';
import { ICalendarEvent } from '../models/ICalendarModels';
import { IRecurrencePattern } from '../models/IEventModels';

export class RecurrenceUtils {
  /**
   * Parse SharePoint recurrence XML data
   */
  public static parseSharePointRecurrence(recurrenceXml: string): IRecurrencePattern | undefined {
    try {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(recurrenceXml, 'text/xml');
      
      const recurrenceElement = xmlDoc.getElementsByTagName('recurrence')[0];
      if (!recurrenceElement) return undefined;

      const rule = recurrenceElement.getElementsByTagName('rule')[0];
      if (!rule) return undefined;

      const repeat = rule.getElementsByTagName('repeat')[0];
      const windowEnd = rule.getElementsByTagName('windowEnd')[0];

      if (!repeat) return undefined;

      const pattern: IRecurrencePattern = {
        type: 'custom',
        interval: 1
      };

      // Parse recurrence type and interval
      if (repeat.getElementsByTagName('daily')[0]) {
        pattern.type = 'daily';
        const dailyFreq = repeat.getElementsByTagName('daily')[0].getAttribute('dayFrequency');
        pattern.interval = dailyFreq ? parseInt(dailyFreq, 10) : 1;
      } else if (repeat.getElementsByTagName('weekly')[0]) {
        pattern.type = 'weekly';
        const weeklyElement = repeat.getElementsByTagName('weekly')[0];
        const weekFreq = weeklyElement.getAttribute('weekFrequency');
        pattern.interval = weekFreq ? parseInt(weekFreq, 10) : 1;

        // Parse days of week
        const daysOfWeek: number[] = [];
        ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'].forEach((day, index) => {
          if (weeklyElement.getAttribute(day) === 'TRUE') {
            daysOfWeek.push(index);
          }
        });
        pattern.daysOfWeek = daysOfWeek;
      } else if (repeat.getElementsByTagName('monthly')[0]) {
        pattern.type = 'monthly';
        const monthlyElement = repeat.getElementsByTagName('monthly')[0];
        const monthFreq = monthlyElement.getAttribute('monthFrequency');
        pattern.interval = monthFreq ? parseInt(monthFreq, 10) : 1;

        const day = monthlyElement.getAttribute('day');
        if (day) {
          pattern.dayOfMonth = parseInt(day, 10);
        }
      } else if (repeat.getElementsByTagName('yearly')[0]) {
        pattern.type = 'yearly';
        const yearlyElement = repeat.getElementsByTagName('yearly')[0];
        const yearFreq = yearlyElement.getAttribute('yearFrequency');
        pattern.interval = yearFreq ? parseInt(yearFreq, 10) : 1;

        const month = yearlyElement.getAttribute('month');
        const day = yearlyElement.getAttribute('day');
        if (month) pattern.monthOfYear = parseInt(month, 10);
        if (day) pattern.dayOfMonth = parseInt(day, 10);
      }

      // Parse end date
      if (windowEnd) {
        const endDateStr = windowEnd.textContent;
        if (endDateStr) {
          pattern.endDate = new Date(endDateStr);
        }
      }

      return pattern;
    } catch (error) {
      console.warn('Error parsing SharePoint recurrence:', error);
      return undefined;
    }
  }

  /**
   * Parse Microsoft Graph recurrence data
   */
  public static parseGraphRecurrence(graphRecurrence: Record<string, unknown>): IRecurrencePattern | undefined {
    try {
      if (!graphRecurrence || !graphRecurrence.pattern) return undefined;

      const pattern: IRecurrencePattern = {
        type: 'custom',
        interval: (graphRecurrence.pattern as Record<string, unknown>).interval as number || 1
      };

      // Map Graph recurrence type to our type
      switch ((graphRecurrence.pattern as Record<string, unknown>).type) {
        case 'daily': {
          pattern.type = 'daily';
          break;
        }
        case 'weekly': {
          pattern.type = 'weekly';
          pattern.daysOfWeek = this.mapGraphDaysOfWeek((graphRecurrence.pattern as Record<string, unknown>).daysOfWeek as string[]);
          pattern.firstDayOfWeek = this.mapGraphDayOfWeek((graphRecurrence.pattern as Record<string, unknown>).firstDayOfWeek as string);
          break;
        }
        case 'absoluteMonthly': {
          pattern.type = 'monthly';
          pattern.dayOfMonth = (graphRecurrence.pattern as Record<string, unknown>).dayOfMonth as number;
          break;
        }
        case 'relativeMonthly': {
          pattern.type = 'monthly';
          pattern.weekOfMonth = this.mapGraphWeekOfMonth((graphRecurrence.pattern as Record<string, unknown>).index as string);
          pattern.daysOfWeek = this.mapGraphDaysOfWeek((graphRecurrence.pattern as Record<string, unknown>).daysOfWeek as string[]);
          break;
        }
        case 'absoluteYearly': {
          pattern.type = 'yearly';
          pattern.dayOfMonth = (graphRecurrence.pattern as Record<string, unknown>).dayOfMonth as number;
          pattern.monthOfYear = (graphRecurrence.pattern as Record<string, unknown>).month as number;
          break;
        }
        case 'relativeYearly': {
          pattern.type = 'yearly';
          pattern.weekOfMonth = this.mapGraphWeekOfMonth((graphRecurrence.pattern as Record<string, unknown>).index as string);
          pattern.daysOfWeek = this.mapGraphDaysOfWeek((graphRecurrence.pattern as Record<string, unknown>).daysOfWeek as string[]);
          pattern.monthOfYear = (graphRecurrence.pattern as Record<string, unknown>).month as number;
          break;
        }
      }

      // Parse range
      if (graphRecurrence.range) {
        const range = graphRecurrence.range as Record<string, unknown>;
        switch (range.type) {
          case 'endDate': {
            pattern.endDate = new Date(range.endDate as string);
            break;
          }
          case 'numbered': {
            pattern.occurrences = range.numberOfOccurrences as number;
            break;
          }
        }
      }

      return pattern;
    } catch (error) {
      console.warn('Error parsing Graph recurrence:', error);
      return undefined;
    }
  }

  /**
   * Generate recurring events from base event and pattern
   */
  public static generateRecurringEvents(
    baseEvent: ICalendarEvent,
    pattern: IRecurrencePattern,
    startDate: Date,
    endDate: Date,
    maxOccurrences: number = 100
  ): ICalendarEvent[] {
    const events: ICalendarEvent[] = [];
    const eventDuration = baseEvent.end.getTime() - baseEvent.start.getTime();
    
    let currentDate = moment(baseEvent.start);
    let occurrenceCount = 0;

    // Ensure we don't generate events before the requested start date
    if (currentDate.isBefore(startDate)) {
      currentDate = moment(startDate);
      // Adjust to next valid occurrence
      currentDate = this.getNextOccurrence(currentDate, pattern, moment(baseEvent.start));
    }

    while (
      currentDate.isBefore(endDate) &&
      occurrenceCount < maxOccurrences &&
      (!pattern.endDate || currentDate.isBefore(pattern.endDate)) &&
      (!pattern.occurrences || occurrenceCount < pattern.occurrences)
    ) {
      // Create recurring event instance
      const recurringEvent: ICalendarEvent = {
        ...baseEvent,
        id: `${baseEvent.id}_recur_${occurrenceCount}`,
        start: currentDate.toDate(),
        end: new Date(currentDate.toDate().getTime() + eventDuration)
      };

      events.push(recurringEvent);
      occurrenceCount++;

      // Calculate next occurrence
      currentDate = this.getNextOccurrence(currentDate, pattern);
    }

    return events;
  }

  /**
   * Get next occurrence based on recurrence pattern
   */
  private static getNextOccurrence(
    currentDate: moment.Moment,
    pattern: IRecurrencePattern,
    baseDate?: moment.Moment
  ): moment.Moment {
    const next = currentDate.clone();

    switch (pattern.type) {
      case 'daily': {
        next.add(pattern.interval, 'days');
        break;
      }
      case 'weekly': {
        if (pattern.daysOfWeek && pattern.daysOfWeek.length > 0) {
          // Find next day of week in the pattern
          const currentDayOfWeek = next.day();
          const sortedDays = [...pattern.daysOfWeek].sort((a, b) => a - b);
          
          const nextDay = sortedDays.find(day => day > currentDayOfWeek);
          
          if (nextDay !== undefined) {
            // Next occurrence is in the same week
            next.day(nextDay);
          } else {
            // Next occurrence is in the next week cycle
            next.add(pattern.interval, 'weeks');
            next.day(sortedDays[0]);
          }
        } else {
          next.add(pattern.interval, 'weeks');
        }
        break;
      }
      case 'monthly': {
        if (pattern.dayOfMonth) {
          next.add(pattern.interval, 'months');
          next.date(pattern.dayOfMonth);
          
          // Handle months with fewer days
          if (next.date() !== pattern.dayOfMonth) {
            next.date(0); // Go to last day of previous month
          }
        } else {
          next.add(pattern.interval, 'months');
        }
        break;
      }
      case 'yearly': {
        next.add(pattern.interval, 'years');
        if (pattern.monthOfYear) {
          next.month(pattern.monthOfYear - 1); // moment months are 0-based
        }
        if (pattern.dayOfMonth) {
          next.date(pattern.dayOfMonth);
        }
        break;
      }
      case 'weekdays': {
        // Skip weekends
        do {
          next.add(1, 'day');
        } while (next.day() === 0 || next.day() === 6); // Sunday = 0, Saturday = 6
        break;
      }
      default: {
        next.add(pattern.interval, 'days');
        break;
      }
    }

    return next;
  }

  /**
   * Map Graph API days of week to our format
   */
  private static mapGraphDaysOfWeek(graphDays: string[]): number[] {
    if (!graphDays) return [];
    
    const dayMap: { [key: string]: number } = {
      'sunday': 0,
      'monday': 1,
      'tuesday': 2,
      'wednesday': 3,
      'thursday': 4,
      'friday': 5,
      'saturday': 6
    };

    return graphDays.map(day => dayMap[day.toLowerCase()]).filter(d => d !== undefined);
  }

  /**
   * Map Graph API day of week to our format
   */
  private static mapGraphDayOfWeek(graphDay: string): number {
    if (!graphDay) return 0;
    
    const dayMap: { [key: string]: number } = {
      'sunday': 0,
      'monday': 1,
      'tuesday': 2,
      'wednesday': 3,
      'thursday': 4,
      'friday': 5,
      'saturday': 6
    };

    return dayMap[graphDay?.toLowerCase()] || 0;
  }

  /**
   * Map Graph API week of month to our format
   */
  private static mapGraphWeekOfMonth(graphIndex: string): number {
    if (!graphIndex) return 1;
    
    const indexMap: { [key: string]: number } = {
      'first': 1,
      'second': 2,
      'third': 3,
      'fourth': 4,
      'last': -1
    };

    return indexMap[graphIndex?.toLowerCase()] || 1;
  }

  /**
   * Check if an event is part of a recurring series
   */
  public static isRecurringEvent(event: ICalendarEvent): boolean {
    return event.isRecurring || !!event.masterSeriesId;
  }

  /**
   * Check if an event is an exception to a recurring series
   */
  public static isRecurringException(event: ICalendarEvent): boolean {
    return !!event.isException || !!event.masterSeriesId;
  }

  /**
   * Find the master event for a recurring series
   */
  public static findMasterEvent(events: ICalendarEvent[], seriesId: string): ICalendarEvent | undefined {
    return events.find(event => 
      event.id === seriesId && 
      event.isRecurring && 
      !event.isException
    );
  }

  /**
   * Find all instances of a recurring series
   */
  public static findSeriesInstances(events: ICalendarEvent[], seriesId: string): ICalendarEvent[] {
    return events.filter(event => 
      event.id === seriesId || 
      event.masterSeriesId === seriesId
    );
  }

  /**
   * Generate recurrence description text
   */
  public static getRecurrenceDescription(pattern: IRecurrencePattern): string {
    if (!pattern) return '';

    const intervalText = pattern.interval > 1 ? ` every ${pattern.interval}` : '';
    
    switch (pattern.type) {
      case 'daily': {
        return `Daily${intervalText} day${pattern.interval > 1 ? 's' : ''}`;
      }
      case 'weekly': {
        let weeklyText = `Weekly${intervalText} week${pattern.interval > 1 ? 's' : ''}`;
        if (pattern.daysOfWeek && pattern.daysOfWeek.length > 0) {
          const dayNames = pattern.daysOfWeek.map(day => moment().day(day).format('dddd'));
          weeklyText += ` on ${dayNames.join(', ')}`;
        }
        return weeklyText;
      }
      case 'monthly': {
        let monthlyText = `Monthly${intervalText} month${pattern.interval > 1 ? 's' : ''}`;
        if (pattern.dayOfMonth) {
          monthlyText += ` on day ${pattern.dayOfMonth}`;
        }
        return monthlyText;
      }
      case 'yearly': {
        let yearlyText = `Yearly${intervalText} year${pattern.interval > 1 ? 's' : ''}`;
        if (pattern.monthOfYear && pattern.dayOfMonth) {
          const monthName = moment().month(pattern.monthOfYear - 1).format('MMMM');
          yearlyText += ` on ${monthName} ${pattern.dayOfMonth}`;
        }
        return yearlyText;
      }
      case 'weekdays': {
        return 'Every weekday (Monday to Friday)';
      }
      default: {
        return 'Custom recurrence pattern';
      }
    }
  }

  /**
   * Get next occurrence date for a recurring event
   */
  public static getNextOccurrenceDate(event: ICalendarEvent, pattern: IRecurrencePattern): Date | undefined {
    if (!pattern) return undefined;

    const nextOccurrence = this.getNextOccurrence(moment(event.start), pattern);
    return nextOccurrence.toDate();
  }

  /**
   * Check if a date matches the recurrence pattern
   */
  public static dateMatchesPattern(date: Date, pattern: IRecurrencePattern, baseDate: Date): boolean {
    const momentDate = moment(date);
    const momentBase = moment(baseDate);

    switch (pattern.type) {
      case 'daily': {
        const daysDiff = momentDate.diff(momentBase, 'days');
        return daysDiff >= 0 && daysDiff % pattern.interval === 0;
      }
      case 'weekly': {
        const weeksDiff = momentDate.diff(momentBase, 'weeks');
        const dayOfWeek = momentDate.day();
        return (
          weeksDiff >= 0 &&
          weeksDiff % pattern.interval === 0 &&
          (!pattern.daysOfWeek || pattern.daysOfWeek.includes(dayOfWeek))
        );
      }
      case 'monthly': {
        const monthsDiff = momentDate.diff(momentBase, 'months');
        return (
          monthsDiff >= 0 &&
          monthsDiff % pattern.interval === 0 &&
          (!pattern.dayOfMonth || momentDate.date() === pattern.dayOfMonth)
        );
      }
      case 'yearly': {
        const yearsDiff = momentDate.diff(momentBase, 'years');
        return (
          yearsDiff >= 0 &&
          yearsDiff % pattern.interval === 0 &&
          (!pattern.monthOfYear || momentDate.month() + 1 === pattern.monthOfYear) &&
          (!pattern.dayOfMonth || momentDate.date() === pattern.dayOfMonth)
        );
      }
      case 'weekdays': {
        const day = momentDate.day();
        return day >= 1 && day <= 5; // Monday to Friday
      }
      default: {
        return false;
      }
    }
  }

  /**
   * Calculate total occurrences for a recurrence pattern
   */
  public static calculateTotalOccurrences(
    pattern: IRecurrencePattern,
    startDate: Date,
    endDate?: Date
  ): number {
    if (pattern.occurrences) {
      return pattern.occurrences;
    }

    if (!endDate && !pattern.endDate) {
      return Number.POSITIVE_INFINITY; // Infinite recurrence
    }

    const finalEndDate = pattern.endDate || endDate;
    if (!finalEndDate) return 0;

    const start = moment(startDate);
    const end = moment(finalEndDate);
    
    switch (pattern.type) {
      case 'daily': {
        return Math.floor(end.diff(start, 'days') / pattern.interval) + 1;
      }
      case 'weekly': {
        return Math.floor(end.diff(start, 'weeks') / pattern.interval) + 1;
      }
      case 'monthly': {
        return Math.floor(end.diff(start, 'months') / pattern.interval) + 1;
      }
      case 'yearly': {
        return Math.floor(end.diff(start, 'years') / pattern.interval) + 1;
      }
      default: {
        return Math.floor(end.diff(start, 'days')) + 1;
      }
    }
  }

  /**
   * Validate recurrence pattern
   */
  public static validateRecurrencePattern(pattern: IRecurrencePattern): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    if (!pattern) {
      errors.push('Recurrence pattern is required');
      return { isValid: false, errors };
    }

    // Validate interval
    if (!pattern.interval || pattern.interval < 1) {
      errors.push('Recurrence interval must be at least 1');
    }

    if (pattern.interval > 999) {
      errors.push('Recurrence interval cannot exceed 999');
    }

    // Validate type-specific properties
    switch (pattern.type) {
      case 'weekly': {
        if (pattern.daysOfWeek) {
          if (pattern.daysOfWeek.length === 0) {
            errors.push('Weekly recurrence must specify at least one day of the week');
          }
          
          const invalidDays = pattern.daysOfWeek.filter(day => day < 0 || day > 6);
          if (invalidDays.length > 0) {
            errors.push('Days of week must be between 0 (Sunday) and 6 (Saturday)');
          }
        }
        break;
      }
      case 'monthly': {
        if (pattern.dayOfMonth && (pattern.dayOfMonth < 1 || pattern.dayOfMonth > 31)) {
          errors.push('Day of month must be between 1 and 31');
        }
        if (pattern.weekOfMonth && (pattern.weekOfMonth < -1 || pattern.weekOfMonth === 0 || pattern.weekOfMonth > 4)) {
          errors.push('Week of month must be between 1-4 or -1 (last week)');
        }
        break;
      }
      case 'yearly': {
        if (pattern.monthOfYear && (pattern.monthOfYear < 1 || pattern.monthOfYear > 12)) {
          errors.push('Month of year must be between 1 and 12');
        }
        if (pattern.dayOfMonth && (pattern.dayOfMonth < 1 || pattern.dayOfMonth > 31)) {
          errors.push('Day of month must be between 1 and 31');
        }
        break;
      }
    }

    // Validate end conditions
    if (pattern.endDate && pattern.occurrences) {
      errors.push('Cannot specify both end date and number of occurrences');
    }

    if (pattern.occurrences && pattern.occurrences < 1) {
      errors.push('Number of occurrences must be at least 1');
    }

    if (pattern.occurrences && pattern.occurrences > 999) {
      errors.push('Number of occurrences cannot exceed 999');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  /**
   * Create a simple recurrence pattern
   */
  public static createSimplePattern(
    type: 'daily' | 'weekly' | 'monthly' | 'yearly',
    interval: number = 1,
    endDate?: Date,
    occurrences?: number
  ): IRecurrencePattern {
    const pattern: IRecurrencePattern = {
      type,
      interval
    };

    if (endDate) {
      pattern.endDate = endDate;
    }

    if (occurrences) {
      pattern.occurrences = occurrences;
    }

    return pattern;
  }

  /**
   * Get human-readable recurrence summary
   */
  public static getSummary(pattern: IRecurrencePattern): string {
    const description = this.getRecurrenceDescription(pattern);
    let summary = description;

    if (pattern.endDate) {
      summary += ` until ${moment(pattern.endDate).format('MMMM D, YYYY')}`;
    } else if (pattern.occurrences) {
      summary += ` for ${pattern.occurrences} occurrence${pattern.occurrences > 1 ? 's' : ''}`;
    }

    return summary;
  }

  /**
   * Clone a recurrence pattern
   */
  public static clone(pattern: IRecurrencePattern): IRecurrencePattern {
    return {
      ...pattern,
      daysOfWeek: pattern.daysOfWeek ? [...pattern.daysOfWeek] : undefined,
      endDate: pattern.endDate ? new Date(pattern.endDate) : undefined
    };
  }
}