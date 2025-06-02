import * as React from 'react';
import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  CommandBar,
  ICommandBarItemProps,
  SearchBox,
  Pivot,
  PivotItem,
  Panel,
  PanelType,
  Icon,
  mergeStyles,
  useTheme,
  IStackTokens,
  ITheme
} from '@fluentui/react';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';

import { IMultiCalendarAggregatorProps } from './IMultiCalendarAggregatorProps';
import { CalendarService } from '../services/CalendarService';
import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';
import { CalendarSourcesPanel } from './CalendarSourcesPanel';
import { AgendaView } from './AgendaView';
import { TimelineView } from './TimelineView';
import { EventDetailsPanel } from './EventDetailsPanel';

const localizer = momentLocalizer(moment);

const stackTokens: IStackTokens = { childrenGap: 15 };

// Error Boundary Component
interface IErrorBoundaryState {
  hasError: boolean;
  error?: Error;
}

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  IErrorBoundaryState
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false };
  }

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { hasError: true, error };
  }

  public override componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error('Calendar component error:', error, errorInfo);
  }

  public override render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          <Text>
            Something went wrong with the calendar component. Please refresh the page.
            {this.state.error && (
              <details style={{ marginTop: '8px' }}>
                <summary>Error details</summary>
                <pre style={{ fontSize: '12px', marginTop: '4px' }}>
                  {this.state.error.message}
                </pre>
              </details>
            )}
          </Text>
        </MessageBar>
      );
    }

    return this.props.children;
  }
}

const MultiCalendarAggregator: React.FC<IMultiCalendarAggregatorProps> = (props) => {
  const theme: ITheme = useTheme();
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [filteredEvents, setFilteredEvents] = useState<ICalendarEvent[]>([]);
  const [calendarSources, setCalendarSources] = useState<ICalendarSource[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(null);
  const [showEventDetails, setShowEventDetails] = useState<boolean>(false);
  const [showCalendarSources, setShowCalendarSources] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentView, setCurrentView] = useState<string>(props.viewType);
  const [currentDate, setCurrentDate] = useState<Date>(new Date());

  // Memoize calendar service to prevent recreation
  const calendarService = useMemo(() => {
    try {
      return new CalendarService(props.context);
    } catch (error) {
      console.error('Failed to create calendar service:', error);
      return null;
    }
  }, [props.context]);

  // Container styles - memoized to prevent recreation
  const containerStyles = useMemo(() => mergeStyles({
    padding: '20px',
    backgroundColor: theme.palette.neutralLighterAlt,
    borderRadius: '8px',
    minHeight: '600px',
    boxShadow: theme.effects.elevation4,
    border: `1px solid ${theme.palette.neutralLight}`,
  }), [theme]);

  const headerStyles = useMemo(() => mergeStyles({
    backgroundColor: theme.palette.white,
    padding: '16px 20px',
    borderRadius: '8px',
    marginBottom: '16px',
    boxShadow: theme.effects.elevation4,
    border: `1px solid ${theme.palette.neutralLight}`,
  }), [theme]);

  // Load calendar data with proper error handling
  const loadCalendarData = useCallback(async () => {
    if (!calendarService) {
      setError('Calendar service initialization failed');
      setLoading(false);
      return;
    }

    try {
      setLoading(true);
      setError('');

      // Load calendar sources with timeout
      const sourcesPromise = calendarService.getCalendarSources(props.useGraphAPI);
      const timeoutPromise = new Promise<never>((_, reject) => {
        setTimeout(() => reject(new Error('Request timeout')), 30000);
      });

      const sources = await Promise.race([sourcesPromise, timeoutPromise]);
      setCalendarSources(sources);

      // Load events from selected calendars
      const selectedSources = sources.filter(s => 
        props.selectedCalendars.length === 0 || props.selectedCalendars.includes(s.id)
      );

      if (selectedSources.length === 0) {
        setEvents([]);
        setFilteredEvents([]);
        setLoading(false);
        return;
      }

      const eventsPromise = calendarService.getEventsFromSources(selectedSources, props.maxEvents);
      const allEvents = await Promise.race([eventsPromise, timeoutPromise]);

      // Sort events by start date
      allEvents.sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());

      setEvents(allEvents);
      setFilteredEvents(allEvents);
    } catch (err) {
      console.error('Error loading calendar data:', err);
      const errorMessage = err instanceof Error ? err.message : 'Unknown error occurred';
      setError(`Failed to load calendar data: ${errorMessage}`);
      setEvents([]);
      setFilteredEvents([]);
    } finally {
      setLoading(false);
    }
  }, [calendarService, props.selectedCalendars, props.useGraphAPI, props.maxEvents]);

  // Filter events based on search query - debounced
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      if (!searchQuery.trim()) {
        setFilteredEvents(events);
        return;
      }

      const query = searchQuery.toLowerCase();
      const filtered = events.filter(event =>
        event.title.toLowerCase().includes(query) ||
        event.description?.toLowerCase().includes(query) ||
        event.location?.toLowerCase().includes(query)
      );

      setFilteredEvents(filtered);
    }, 300); // Debounce search

    return () => clearTimeout(timeoutId);
  }, [events, searchQuery]);

  // Initial load and auto-refresh with proper cleanup
  useEffect(() => {
    let mounted = true;
    let refreshInterval: number | undefined;

    const initialLoad = async (): Promise<void> => {
      if (mounted) {
        await loadCalendarData();
      }
    };

    // Initial load
    void initialLoad();

    // Set up auto-refresh
    if (props.refreshInterval > 0) {
      refreshInterval = window.setInterval(() => {
        if (mounted) {
          void loadCalendarData();
        }
      }, props.refreshInterval * 60 * 1000);
    }

    // Cleanup
    return (): void => {
      mounted = false;
      if (refreshInterval) {
        window.clearInterval(refreshInterval);
      }
    };
  }, [loadCalendarData, props.refreshInterval]);

  // Command bar items - memoized
  const commandBarItems: ICommandBarItemProps[] = useMemo(() => [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      disabled: loading,
      onClick: (): void => {
        void loadCalendarData();
      },
    },
    {
      key: 'sources',
      text: 'Manage Sources',
      iconProps: { iconName: 'Calendar' },
      onClick: (): void => setShowCalendarSources(true),
    },
    {
      key: 'today',
      text: 'Today',
      iconProps: { iconName: 'DateTime' },
      onClick: (): void => setCurrentDate(new Date()),
    },
  ], [loading, loadCalendarData]);

  const commandBarFarItems: ICommandBarItemProps[] = useMemo(() => [
    {
      key: 'search',
      onRender: () => (
        <SearchBox
          placeholder="Search events..."
          value={searchQuery}
          onChange={(_, value) => setSearchQuery(value || '')}
          styles={{
            root: { width: '250px', marginTop: '4px' }
          }}
          disabled={loading}
        />
      ),
    },
  ], [searchQuery, loading]);

  // Event handlers
  const handleEventSelect = useCallback((event: ICalendarEvent): void => {
    setSelectedEvent(event);
    setShowEventDetails(true);
  }, []);

  const handleNavigate = useCallback((date: Date): void => {
    setCurrentDate(date);
  }, []);

  const handleViewChange = useCallback((view: string): void => {
    setCurrentView(view);
  }, []);

  // Event style getter for color coding - memoized
  const eventStyleGetter = useCallback((event: ICalendarEvent): { style: React.CSSProperties } => {
    if (!props.colorCoding) {
      return {
        style: {
          backgroundColor: theme.palette.themePrimary,
          borderColor: theme.palette.themeDark,
          color: theme.palette.white,
        }
      };
    }

    const source = calendarSources.find(s => s.id === event.calendarId);
    const backgroundColor = source?.color || theme.palette.themePrimary;
    
    return {
      style: {
        backgroundColor,
        borderColor: backgroundColor,
        color: theme.palette.white,
      }
    };
  }, [props.colorCoding, theme, calendarSources]);

  // Render different views - memoized
  const renderCalendarView = useCallback((): React.ReactElement => {
    if (loading) {
      return (
        <div style={{ padding: '40px', textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Loading events..." />
        </div>
      );
    }

    switch (currentView) {
      case 'agenda':
        return (
          <AgendaView
            events={filteredEvents}
            onEventSelect={handleEventSelect}
            calendarSources={calendarSources}
            theme={theme}
          />
        );
      case 'timeline':
        return (
          <TimelineView
            events={filteredEvents}
            onEventSelect={handleEventSelect}
            calendarSources={calendarSources}
            theme={theme}
            currentDate={currentDate}
          />
        );
      case 'month':
      case 'week':
      case 'day':
        return (
          <Calendar
            localizer={localizer}
            events={filteredEvents}
            startAccessor="start"
            endAccessor="end"
            titleAccessor="title"
            style={{ height: '600px' }}
            onSelectEvent={handleEventSelect}
            onNavigate={handleNavigate}
            onView={handleViewChange}
            view={currentView as any}
            date={currentDate}
            eventPropGetter={eventStyleGetter}
            showAllEvents={true}
            popup={true}
            views={{
              month: true,
              week: true,
              day: true,
            }}
            formats={{
              timeGutterFormat: 'HH:mm',
              eventTimeRangeFormat: ({ start, end }) =>
                `${moment(start).format('HH:mm')} - ${moment(end).format('HH:mm')}`,
            }}
          />
        );
      default:
        return (
          <div style={{ padding: '20px', textAlign: 'center' }}>
            <Text variant="large">View not available</Text>
          </div>
        );
    }
  }, [currentView, filteredEvents, handleEventSelect, calendarSources, theme, currentDate, handleNavigate, handleViewChange, eventStyleGetter, loading]);

  // Early return for loading state
  if (loading && events.length === 0) {
    return (
      <ErrorBoundary>
        <div className={containerStyles}>
          <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '400px' } }}>
            <Spinner size={SpinnerSize.large} label="Loading calendar events..." />
          </Stack>
        </div>
      </ErrorBoundary>
    );
  }

  return (
    <ErrorBoundary>
      <div className={containerStyles}>
        {/* Header */}
        <div className={headerStyles}>
          <Stack tokens={stackTokens}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Stack>
                <Text variant="xLarge" styles={{ root: { fontWeight: 600, color: theme.palette.themePrimary } }}>
                  {props.title || 'Multi-Calendar Aggregator'}
                </Text>
                <Text variant="medium" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  {filteredEvents.length} events from {calendarSources.length} calendar sources
                </Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <Icon
                  iconName="Calendar"
                  styles={{
                    root: {
                      fontSize: '24px',
                      color: theme.palette.themePrimary,
                    }
                  }}
                />
              </Stack>
            </Stack>
          </Stack>
        </div>

        {/* Command Bar */}
        <CommandBar
          items={commandBarItems}
          farItems={commandBarFarItems}
          styles={{
            root: {
              backgroundColor: theme.palette.white,
              borderRadius: '6px',
              marginBottom: '16px',
              boxShadow: theme.effects.elevation4,
            }
          }}
        />

        {/* Error Message */}
        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => setError('')}
            dismissButtonAriaLabel="Close"
            styles={{ root: { marginBottom: '16px' } }}
          >
            {error}
          </MessageBar>
        )}

        {/* View Selector */}
        <Pivot
          selectedKey={currentView}
          onLinkClick={(item) => item && setCurrentView(item.props.itemKey!)}
          styles={{
            root: {
              backgroundColor: theme.palette.white,
              borderRadius: '6px',
              marginBottom: '16px',
              boxShadow: theme.effects.elevation4,
              padding: '8px 16px',
            }
          }}
        >
          <PivotItem headerText="Month" itemKey="month" itemIcon="Calendar" />
          <PivotItem headerText="Week" itemKey="week" itemIcon="CalendarWeek" />
          <PivotItem headerText="Day" itemKey="day" itemIcon="CalendarDay" />
          <PivotItem headerText="Agenda" itemKey="agenda" itemIcon="BulletedList" />
          <PivotItem headerText="Timeline" itemKey="timeline" itemIcon="Timeline" />
        </Pivot>

        {/* Calendar View */}
        {renderCalendarView()}

        {/* Event Details Panel */}
        <Panel
          isOpen={showEventDetails}
          onDismiss={() => setShowEventDetails(false)}
          type={PanelType.medium}
          headerText="Event Details"
          closeButtonAriaLabel="Close"
        >
          {selectedEvent && (
            <EventDetailsPanel
              event={selectedEvent}
              calendarSource={calendarSources.find(s => s.id === selectedEvent.calendarId)}
              onClose={() => setShowEventDetails(false)}
            />
          )}
        </Panel>

        {/* Calendar Sources Panel */}
        <Panel
          isOpen={showCalendarSources}
          onDismiss={() => setShowCalendarSources(false)}
          type={PanelType.medium}
          headerText="Manage Calendar Sources"
          closeButtonAriaLabel="Close"
        >
          <CalendarSourcesPanel
            sources={calendarSources}
            selectedSources={props.selectedCalendars}
            onSourcesChange={(selected) => {
              // Update web part properties
              console.log('Selected sources changed:', selected);
            }}
            onClose={() => setShowCalendarSources(false)}
          />
        </Panel>
      </div>
    </ErrorBoundary>
  );
};

export default MultiCalendarAggregator;