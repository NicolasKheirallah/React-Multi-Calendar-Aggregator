import * as React from 'react';
import { useState, useMemo, useRef, useEffect } from 'react';
import {
  Stack,
  Text,
  ActionButton,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  ScrollablePane,
  Sticky,
  StickyPositionType,
  ITheme
} from '@fluentui/react';
import moment from 'moment';

import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';

export interface ITimelineViewProps {
  events: ICalendarEvent[];
  onEventSelect: (event: ICalendarEvent) => void;
  calendarSources: ICalendarSource[];
  theme: ITheme;
  currentDate: Date;
  timeRange?: 'day' | 'week' | 'month';
}

export const TimelineView: React.FC<ITimelineViewProps> = ({
  events,
  onEventSelect,
  calendarSources,
  theme,
  currentDate,
  timeRange = 'week'
}) => {
  const [selectedTimeRange, setSelectedTimeRange] = useState<string>(timeRange);
  const [currentViewDate, setCurrentViewDate] = useState<Date>(currentDate);
  const timelineRef = useRef<HTMLDivElement>(null);

  const containerStyles = mergeStyles({
    height: '600px',
    backgroundColor: theme.palette.neutralLighterAlt,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: '8px',
    overflow: 'hidden',
    position: 'relative'
  });

  const headerStyles = mergeStyles({
    backgroundColor: theme.palette.white,
    borderBottom: `1px solid ${theme.palette.neutralLight}`,
    padding: '12px 16px',
    zIndex: 100
  });

  const timelineContainerStyles = mergeStyles({
    height: 'calc(100% - 60px)',
    overflow: 'auto',
    position: 'relative'
  });

  const timelineGridStyles = mergeStyles({
    position: 'relative',
    minHeight: '500px',
    backgroundColor: theme.palette.white
  });

  const timeSlotStyles = mergeStyles({
    borderRight: `1px solid ${theme.palette.neutralLighter}`,
    borderBottom: `1px solid ${theme.palette.neutralLighter}`,
    minHeight: '60px',
    position: 'relative',
    '&:hover': {
      backgroundColor: theme.palette.neutralLighterAlt
    }
  });

  const eventBarStyles = mergeStyles({
    position: 'absolute',
    borderRadius: '4px',
    padding: '4px 8px',
    color: theme.palette.white,
    fontSize: '12px',
    fontWeight: 500,
    cursor: 'pointer',
    transition: 'all 0.2s ease',
    boxShadow: theme.effects.elevation4,
    '&:hover': {
      transform: 'translateY(-1px)',
      boxShadow: theme.effects.elevation8,
      zIndex: 50
    }
  });

  const timeRangeOptions: IDropdownOption[] = [
    { key: 'day', text: 'Day View' },
    { key: 'week', text: 'Week View' },
    { key: 'month', text: 'Month View' }
  ];

  // Calculate date range based on selected time range
  const getDateRange = (): { start: Date; end: Date; dates: Date[] } => {
    const start = moment(currentViewDate);
    let end: moment.Moment;
    const dates: Date[] = [];

    switch (selectedTimeRange) {
      case 'day': {
        end = start.clone().endOf('day');
        dates.push(start.toDate());
        break;
      }
      case 'month': {
        start.startOf('month');
        end = start.clone().endOf('month');
        const current = start.clone();
        while (current.isSameOrBefore(end, 'day')) {
          dates.push(current.toDate());
          current.add(1, 'day');
        }
        break;
      }
      case 'week':
      default: {
        start.startOf('week');
        end = start.clone().endOf('week');
        for (let i = 0; i < 7; i++) {
          dates.push(start.clone().add(i, 'days').toDate());
        }
        break;
      }
    }

    return {
      start: start.toDate(),
      end: end.toDate(),
      dates
    };
  };

  // Generate time slots (hours)
  const generateTimeSlots = (): string[] => {
    const slots: string[] = [];
    for (let hour = 0; hour < 24; hour++) {
      slots.push(moment().hour(hour).minute(0).format('HH:mm'));
    }
    return slots;
  };

  const { start: rangeStart, end: rangeEnd, dates } = getDateRange();
  const timeSlots = generateTimeSlots();

  // Filter events for current date range
  const filteredEvents = useMemo(() => {
    return events.filter(event => {
      const eventStart = moment(event.start);
      return eventStart.isBetween(rangeStart, rangeEnd, 'day', '[]');
    });
  }, [events, rangeStart, rangeEnd]);

  // Calculate event positioning
  const getEventPosition = (event: ICalendarEvent, dateIndex: number, columnWidth: number): {
    top: string;
    left: string;
    width: string;
    height: string;
  } => {
    const eventStart = moment(event.start);
    const eventEnd = moment(event.end);
    const dayStart = moment(dates[dateIndex]).startOf('day');
    
    // Calculate top position (based on time)
    const startMinutes = eventStart.diff(dayStart, 'minutes');
    const top = Math.max(0, (startMinutes / 60) * 60); // 60px per hour
    
    // Calculate height (duration)
    const durationMinutes = eventEnd.diff(eventStart, 'minutes');
    const height = Math.max(20, (durationMinutes / 60) * 60);
    
    // Calculate left position and width
    const left = dateIndex * columnWidth;
    const width = columnWidth - 8; // 4px margin on each side
    
    return {
      top: `${top}px`,
      left: `${left + 4}px`,
      width: `${width}px`,
      height: `${height}px`
    };
  };

  const navigate = (direction: 'prev' | 'next'): void => {
    const newDate = moment(currentViewDate);
    
    switch (selectedTimeRange) {
      case 'day': {
        newDate.add(direction === 'next' ? 1 : -1, 'day');
        break;
      }
      case 'month': {
        newDate.add(direction === 'next' ? 1 : -1, 'month');
        break;
      }
      case 'week':
      default: {
        newDate.add(direction === 'next' ? 1 : -1, 'week');
        break;
      }
    }
    
    setCurrentViewDate(newDate.toDate());
  };

  const goToToday = (): void => {
    setCurrentViewDate(new Date());
  };

  // Auto-scroll to current time on mount
  useEffect(() => {
    if (timelineRef.current && selectedTimeRange === 'day') {
      const currentHour = moment().hour();
      const scrollTop = currentHour * 60; // 60px per hour
      timelineRef.current.scrollTop = scrollTop;
    }
  }, [selectedTimeRange, currentViewDate]);

  const renderTimelineGrid = (): React.ReactElement => {
    const columnWidth = selectedTimeRange === 'day' ? 800 : 100 / dates.length;
    const isPercentage = selectedTimeRange !== 'day';

    return (
      <div className={timelineGridStyles} style={{ minWidth: selectedTimeRange === 'day' ? '800px' : '100%' }}>
        {/* Time slots (rows) */}
        {timeSlots.map((timeSlot, timeIndex) => (
          <div key={timeSlot} style={{ display: 'flex', height: '60px' }}>
            {/* Time label */}
            <div
              style={{
                width: '80px',
                minWidth: '80px',
                borderRight: `2px solid ${theme.palette.neutralLight}`,
                borderBottom: `1px solid ${theme.palette.neutralLighter}`,
                padding: '8px',
                backgroundColor: theme.palette.neutralLighterAlt,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: 500,
                fontSize: '12px',
                color: theme.palette.neutralSecondary
              }}
            >
              {timeSlot}
            </div>
            
            {/* Date columns */}
            {dates.map((date, dateIndex) => (
              <div
                key={`${timeSlot}-${dateIndex}`}
                className={timeSlotStyles}
                style={{
                  width: isPercentage ? `${columnWidth}%` : `${columnWidth}px`,
                  backgroundColor: moment(date).isSame(moment(), 'day') ? theme.palette.themeLighter : 'transparent'
                }}
              />
            ))}
          </div>
        ))}

        {/* Events overlay */}
        {filteredEvents.map((event, eventIndex) => {
          const eventDate = moment(event.start);
          const dateIndex = dates.findIndex(date => moment(date).isSame(eventDate, 'day'));
          
          if (dateIndex === -1) return null;

          const position = getEventPosition(event, dateIndex, isPercentage ? (window.innerWidth - 100) * (columnWidth / 100) : columnWidth);
          return (
            <div
              key={event.id}
              className={eventBarStyles}
              style={{
                ...position,
                backgroundColor: event.color || theme.palette.themePrimary,
                marginLeft: '80px', // Account for time column
                zIndex: 10 + eventIndex
              }}
              onClick={() => onEventSelect(event)}
              title={`${event.title}\n${moment(event.start).format('HH:mm')} - ${moment(event.end).format('HH:mm')}\n${event.location || ''}`}
            >
              <Stack tokens={{ childrenGap: 2 }}>
                <Text
                  variant="xSmall"
                  styles={{
                    root: {
                      color: theme.palette.white,
                      fontWeight: 600,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap'
                    }
                  }}
                >
                  {event.title}
                </Text>
                
                {!event.isAllDay && (
                  <Text
                    variant="xSmall"
                    styles={{
                      root: {
                        color: theme.palette.white,
                        opacity: 0.9,
                        fontSize: '10px'
                      }
                    }}
                  >
                    {moment(event.start).format('HH:mm')} - {moment(event.end).format('HH:mm')}
                  </Text>
                )}
                
                {event.location && position.height.replace('px', '') > '40' && (
                  <Text
                    variant="xSmall"
                    styles={{
                      root: {
                        color: theme.palette.white,
                        opacity: 0.8,
                        fontSize: '10px',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap'
                      }
                    }}
                  >
                    📍 {event.location}
                  </Text>
                )}
              </Stack>
            </div>
          );
        })}

        {/* Current time indicator (for day view) */}
        {selectedTimeRange === 'day' && moment(currentViewDate).isSame(moment(), 'day') && (
          <div
            style={{
              position: 'absolute',
              left: '80px',
              right: '0',
              top: `${(moment().hour() + moment().minute() / 60) * 60}px`,
              height: '2px',
              backgroundColor: theme.palette.red,
              zIndex: 100,
              boxShadow: `0 0 4px ${theme.palette.red}`
            }}
          >
            <div
              style={{
                width: '8px',
                height: '8px',
                backgroundColor: theme.palette.red,
                borderRadius: '50%',
                marginLeft: '-4px',
                marginTop: '-3px'
              }}
            />
          </div>
        )}
      </div>
    );
  };

  const getViewTitle = (): string => {
    switch (selectedTimeRange) {
      case 'day':
        return moment(currentViewDate).format('dddd, MMMM Do, YYYY');
      case 'month':
        return moment(currentViewDate).format('MMMM YYYY');
      case 'week':
      default:
        const weekStart = moment(currentViewDate).startOf('week');
        const weekEnd = moment(currentViewDate).endOf('week');
        return `${weekStart.format('MMM D')} - ${weekEnd.format('MMM D, YYYY')}`;
    }
  };

  return (
    <div className={containerStyles}>
      {/* Header */}
      <Sticky stickyPosition={StickyPositionType.Header}>
        <div className={headerStyles}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
              <ActionButton
                iconProps={{ iconName: 'ChevronLeft' }}
                onClick={() => navigate('prev')}
                title="Previous"
              />
              
              <Text variant="large" styles={{ root: { fontWeight: 600, minWidth: '200px', textAlign: 'center' } }}>
                {getViewTitle()}
              </Text>
              
              <ActionButton
                iconProps={{ iconName: 'ChevronRight' }}
                onClick={() => navigate('next')}
                title="Next"
              />
              
              <ActionButton
                iconProps={{ iconName: 'DateTime' }}
                text="Today"
                onClick={goToToday}
              />
            </Stack>

            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
              <Text variant="medium" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {filteredEvents.length} events
              </Text>
              
              <Dropdown
                options={timeRangeOptions}
                selectedKey={selectedTimeRange}
                onChange={(_, option) => option && setSelectedTimeRange(option.key as string)}
                styles={{ root: { minWidth: '120px' } }}
              />
            </Stack>
          </Stack>

          {/* Date headers */}
          {selectedTimeRange !== 'day' && (
            <Stack horizontal style={{ marginTop: '12px' }}>
              <div style={{ width: '80px', minWidth: '80px' }} /> {/* Spacer for time column */}
              {dates.map((date, index) => (
                <div
                  key={index}
                  style={{
                    flex: 1,
                    textAlign: 'center',
                    padding: '8px',
                    borderBottom: `2px solid ${moment(date).isSame(moment(), 'day') ? theme.palette.themePrimary : theme.palette.neutralLight}`,
                    backgroundColor: moment(date).isSame(moment(), 'day') ? theme.palette.themeLighter : 'transparent',
                    fontWeight: moment(date).isSame(moment(), 'day') ? 600 : 400,
                    color: moment(date).isSame(moment(), 'day') ? theme.palette.themePrimary : theme.palette.neutralPrimary
                  }}
                >
                  <Text variant="small" styles={{ root: { fontWeight: 'inherit' } }}>
                    {moment(date).format('ddd')}
                  </Text>
                  <br />
                  <Text variant="medium" styles={{ root: { fontWeight: 'inherit' } }}>
                    {moment(date).format('D')}
                  </Text>
                </div>
              ))}
            </Stack>
          )}
        </div>
      </Sticky>

      {/* Timeline Grid */}
      <div className={timelineContainerStyles} ref={timelineRef}>
        <ScrollablePane>
          {renderTimelineGrid()}
        </ScrollablePane>
      </div>
    </div>
  );
};