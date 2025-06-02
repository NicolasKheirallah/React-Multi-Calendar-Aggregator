import moment from 'moment';
import 'moment-timezone';

export class DateUtils {
  /**
   * Format date for display
   */
  public static formatDate(date: Date, format?: string): string {
    return moment(date).format(format || 'MMMM Do, YYYY');
  }

  /**
   * Format time for display
   */
  public static formatTime(date: Date, format?: string): string {
    return moment(date).format(format || 'h:mm A');
  }

  /**
   * Format date and time for display
   */
  public static formatDateTime(date: Date, format?: string): string {
    return moment(date).format(format || 'MMMM Do, YYYY [at] h:mm A');
  }

  /**
   * Get relative time (e.g., "2 hours ago", "in 3 days")
   */
  public static getRelativeTime(date: Date): string {
    return moment(date).fromNow();
  }

  /**
   * Get duration between two dates
   */
  public static getDuration(start: Date, end: Date): string {
    const duration = moment.duration(moment(end).diff(moment(start)));
    
    const days = Math.floor(duration.asDays());
    const hours = Math.floor(duration.asHours()) % 24;
    const minutes = Math.floor(duration.asMinutes()) % 60;
    
    if (days > 0) {
      return `${days} day${days > 1 ? 's' : ''} ${hours > 0 ? `${hours} hour${hours > 1 ? 's' : ''}` : ''}`.trim();
    } else if (hours > 0) {
      return `${hours} hour${hours > 1 ? 's' : ''} ${minutes > 0 ? `${minutes} minute${minutes > 1 ? 's' : ''}` : ''}`.trim();
    } else {
      return `${minutes} minute${minutes !== 1 ? 's' : ''}`;
    }
  }

  /**
   * Check if date is today
   */
  public static isToday(date: Date): boolean {
    return moment(date).isSame(moment(), 'day');
  }

  /**
   * Check if date is tomorrow
   */
  public static isTomorrow(date: Date): boolean {
    return moment(date).isSame(moment().add(1, 'day'), 'day');
  }

  /**
   * Check if date is this week
   */
  public static isThisWeek(date: Date): boolean {
    return moment(date).isSame(moment(), 'week');
  }

  /**
   * Check if date is next week
   */
  public static isNextWeek(date: Date): boolean {
    return moment(date).isSame(moment().add(1, 'week'), 'week');
  }

  /**
   * Check if date is this month
   */
  public static isThisMonth(date: Date): boolean {
    return moment(date).isSame(moment(), 'month');
  }

  /**
   * Get start of day
   */
  public static getStartOfDay(date: Date): Date {
    return moment(date).startOf('day').toDate();
  }

  /**
   * Get end of day
   */
  public static getEndOfDay(date: Date): Date {
    return moment(date).endOf('day').toDate();
  }

  /**
   * Get start of week
   */
  public static getStartOfWeek(date: Date): Date {
    return moment(date).startOf('week').toDate();
  }

  /**
   * Get end of week
   */
  public static getEndOfWeek(date: Date): Date {
    return moment(date).endOf('week').toDate();
  }

  /**
   * Get start of month
   */
  public static getStartOfMonth(date: Date): Date {
    return moment(date).startOf('month').toDate();
  }

  /**
   * Get end of month
   */
  public static getEndOfMonth(date: Date): Date {
    return moment(date).endOf('month').toDate();
  }

  /**
   * Add time to date
   */
  public static addTime(date: Date, amount: number, unit: moment.unitOfTime.DurationConstructor): Date {
    return moment(date).add(amount, unit).toDate();
  }

  /**
   * Subtract time from date
   */
  public static subtractTime(date: Date, amount: number, unit: moment.unitOfTime.DurationConstructor): Date {
    return moment(date).subtract(amount, unit).toDate();
  }

  /**
   * Get date range for calendar view
   */
  public static getCalendarDateRange(date: Date, viewType: 'month' | 'week' | 'day'): { start: Date; end: Date } {
    const momentDate = moment(date);
    
    switch (viewType) {
      case 'month':
        return {
          start: momentDate.clone().startOf('month').startOf('week').toDate(),
          end: momentDate.clone().endOf('month').endOf('week').toDate()
        };
      case 'week':
        return {
          start: momentDate.clone().startOf('week').toDate(),
          end: momentDate.clone().endOf('week').toDate()
        };
      case 'day':
        return {
          start: momentDate.clone().startOf('day').toDate(),
          end: momentDate.clone().endOf('day').toDate()
        };
      default:
        return {
          start: momentDate.clone().startOf('week').toDate(),
          end: momentDate.clone().endOf('week').toDate()
        };
    }
  }

  /**
   * Get time until event
   */
  public static getTimeUntilEvent(eventDate: Date): string {
    const now = moment();
    const event = moment(eventDate);
    const diff = event.diff(now);
    
    if (diff < 0) {
      return 'Past event';
    }
    
    const duration = moment.duration(diff);
    
    if (duration.asDays() >= 1) {
      return `In ${Math.floor(duration.asDays())} day${Math.floor(duration.asDays()) !== 1 ? 's' : ''}`;
    } else if (duration.asHours() >= 1) {
      return `In ${Math.floor(duration.asHours())} hour${Math.floor(duration.asHours()) !== 1 ? 's' : ''}`;
    } else {
      return `In ${Math.floor(duration.asMinutes())} minute${Math.floor(duration.asMinutes()) !== 1 ? 's' : ''}`;
    }
  }

  /**
   * Check if event is happening now
   */
  public static isEventHappeningNow(startDate: Date, endDate: Date): boolean {
    const now = moment();
    return now.isBetween(moment(startDate), moment(endDate));
  }

  /**
   * Check if event is starting soon (within 15 minutes)
   */
  public static isEventStartingSoon(startDate: Date): boolean {
    const now = moment();
    const start = moment(startDate);
    const diff = start.diff(now, 'minutes');
    return diff > 0 && diff <= 15;
  }

  /**
   * Get friendly date description
   */
  public static getFriendlyDate(date: Date): string {
    const momentDate = moment(date);
    const now = moment();
    
    if (momentDate.isSame(now, 'day')) {
      return 'Today';
    } else if (momentDate.isSame(now.clone().add(1, 'day'), 'day')) {
      return 'Tomorrow';
    } else if (momentDate.isSame(now.clone().subtract(1, 'day'), 'day')) {
      return 'Yesterday';
    } else if (momentDate.isSame(now, 'week')) {
      return momentDate.format('dddd');
    } else if (momentDate.isSame(now.clone().add(1, 'week'), 'week')) {
      return `Next ${momentDate.format('dddd')}`;
    } else if (momentDate.isSame(now, 'year')) {
      return momentDate.format('MMM D');
    } else {
      return momentDate.format('MMM D, YYYY');
    }
  }

  /**
   * Parse SharePoint date string
   */
  public static parseSharePointDate(dateString: string): Date {
    return new Date(dateString);
  }

  /**
   * Format date for SharePoint
   */
  public static formatForSharePoint(date: Date): string {
    return date.toISOString();
  }

  /**
   * Get business days between two dates
   */
  public static getBusinessDaysBetween(start: Date, end: Date): number {
    const startMoment = moment(start);
    const endMoment = moment(end);
    let businessDays = 0;
    
    const current = startMoment.clone();
    while (current.isSameOrBefore(endMoment, 'day')) {
      if (current.day() !== 0 && current.day() !== 6) { // Not Sunday (0) or Saturday (6)
        businessDays++;
      }
      current.add(1, 'day');
    }
    
    return businessDays;
  }

  /**
   * Get timezone offset
   */
  public static getTimezoneOffset(): string {
    return moment().format('Z');
  }

  /**
   * Convert to user's timezone
   */
  public static convertToUserTimezone(date: Date, timezone?: string): Date {
    if (timezone && moment.tz) {
      return moment.tz(date, timezone).toDate();
    }
    return date;
  }
}