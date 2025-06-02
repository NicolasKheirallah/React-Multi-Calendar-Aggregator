import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
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

  const calendarService = new CalendarService(props.context);

  // Container styles
  const containerStyles = mergeStyles({
    padding: '20px',
    backgroundColor: theme.palette.neutralLighterAlt,
    borderRadius: '8px',
    minHeight: '600px',
    boxShadow: theme.effects.elevation4,
    border: `1px solid ${theme.palette.neutralLight}`,
    '& .rbc-calendar': {
      backgroundColor: theme.palette.white,
      borderRadius: '6px',
      overflow: 'hidden',
      boxShadow: theme.effects.elevation8,
      border: `1px solid ${theme.palette.neutralLight}`,
    },
    '& .rbc-header': {
      backgroundColor: theme.palette.themePrimary,
      color: theme.palette.white,
      fontWeight: 600,
      padding: '12px 8px',
      borderBottom: 'none',
    },
    '& .rbc-today': {
      backgroundColor: theme.palette.themeLighter,
    },
    '& .rbc-event': {
      borderRadius: '4px',
      border: 'none',
      fontWeight: 500,
      fontSize: '12px',
      boxShadow: theme.effects.elevation4,
    },
    '& .rbc-month-view': {
      border: 'none',
    },
    '& .rbc-day-bg': {
      borderRight: `1px solid ${theme.palette.neutralLighter}`,
      borderBottom: `1px solid ${theme.palette.neutralLighter}`,
    },
    '& .rbc-time-view .rbc-time-gutter': {
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRight: `1px solid ${theme.palette.neutralLight}`,
    },
    '& .rbc-current-time-indicator': {
      backgroundColor: theme.palette.red,
      height: '2px',
    }
  });

  const headerStyles = mergeStyles({
    backgroundColor: theme.palette.white,
    padding: '16px 20px',
    borderRadius: '8px',
    marginBottom: '16px',
    boxShadow: theme.effects.elevation4,
    border: `1px solid ${theme.palette.neutralLight}`,
  });

  // Load calendar data
  const loadCalendarData = useCallback(async () => {
    try {
      setLoading(true);
      setError('');

      // Load calendar sources
      const sources = await calendarService.getCalendarSources(props.useGraphAPI);
      setCalendarSources(sources);

      // Load events from selected calendars
      const selectedSources = sources.filter(s => 
        props.selectedCalendars.length === 0 || props.selectedCalendars.includes(s.id)
      );

      const allEvents: ICalendarEvent[] = [];
      for (const source of selectedSources) {
        const sourceEvents = await calendarService.getEventsFromSource(source, props.maxEvents);
        allEvents.push(...sourceEvents);
      }

      // Sort events by start date
      allEvents.sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());

      setEvents(allEvents);
      setFilteredEvents(allEvents);
    } catch (err) {
      console.error('Error loading calendar data:', err);
      setError('Failed to load calendar data. Please check your permissions and try again.');
    } finally {
      setLoading(false);
    }
  }, [props.selectedCalendars, props.useGraphAPI, props.maxEvents, calendarService]);

  // Filter events based on search query
  useEffect(() => {
    if (!searchQuery.trim()) {
      setFilteredEvents(events);
      return;
    }

    const filtered = events.filter(event =>
      event.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      event.description?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      event.location?.toLowerCase().includes(searchQuery.toLowerCase())
    );

    setFilteredEvents(filtered);
  }, [events, searchQuery]);

  // Initial load and auto-refresh
  useEffect(() => {
    void loadCalendarData();

    if (props.refreshInterval > 0) {
      const interval = setInterval(() => {
        void loadCalendarData();
      }, props.refreshInterval * 60 * 1000);
      return () => clearInterval(interval);
    }
    
    return (): void => {
      // Empty cleanup function
    };
  }, [loadCalendarData, props.refreshInterval]);

  // Command bar items
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
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
  ];

  const commandBarFarItems: ICommandBarItemProps[] = [
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
        />
      ),
    },
  ];

  // Event handlers
  const handleEventSelect = (event: ICalendarEvent): void => {
    setSelectedEvent(event);
    setShowEventDetails(true);
  };

  const handleNavigate = (date: Date): void => {
    setCurrentDate(date);
  };

  const handleViewChange = (view: string): void => {
    setCurrentView(view);
  };

  // Event style getter for color coding
  const eventStyleGetter = (event: ICalendarEvent): { style: React.CSSProperties } => {
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
  };

  // Render different views
  const renderCalendarView = (): React.ReactElement => {
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
  };

  if (loading) {
    return (
      <div className={containerStyles}>
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '400px' } }}>
          <Spinner size={SpinnerSize.large} label="Loading calendar events..." />
        </Stack>
      </div>
    );
  }

  return (
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
            // This would typically be passed down as a prop
            console.log('Selected sources changed:', selected);
          }}
          onClose={() => setShowCalendarSources(false)}
        />
      </Panel>
    </div>
  );
};

export default MultiCalendarAggregator;