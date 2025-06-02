import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMultiCalendarAggregatorProps {
  title: string;
  selectedCalendars: string[];
  viewType: 'month' | 'week' | 'day' | 'agenda' | 'timeline';
  showWeekends: boolean;
  maxEvents: number;
  refreshInterval: number;
  useGraphAPI: boolean;
  colorCoding: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}