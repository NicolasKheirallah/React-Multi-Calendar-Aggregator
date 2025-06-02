import * as React from 'react';
import { useState, useMemo } from 'react';
import {
  Stack,
  Text,
  Icon,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  IStackTokens,
  Link,
  ITheme
} from '@fluentui/react';
import moment from 'moment';

import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';
import { DateUtils } from '../utils/DateUtils';

export interface IAgendaViewProps {
  events: ICalendarEvent[];
  onEventSelect: (event: ICalendarEvent) => void;
  calendarSources: ICalendarSource[];
  theme: ITheme;
  groupBy?: 'date' | 'calendar' | 'category';
  showDays?: number;
}

const stackTokens: IStackTokens = { childrenGap: 8 };

export const AgendaView: React.FC<IAgendaViewProps> = ({
  events,
  onEventSelect,
  calendarSources,
  theme,
  groupBy = 'date',
  showDays = 14
}) => {
  const [selectedGroupBy, setSelectedGroupBy] = useState<string>(groupBy);
  const [selectedDays, setSelectedDays] = useState<number>(showDays);

  const containerStyles = mergeStyles({
    height: '600px',
    backgroundColor: theme.palette.white,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: '8px',
    overflow: 'hidden'
  });

  const headerStyles = mergeStyles({
    backgroundColor: theme.palette.neutralLighterAlt,
    borderBottom: `1px solid ${theme.palette.neutralLight}`,
    padding: '12px 16px'
  });

  const contentStyles = mergeStyles({
    height: 'calc(100% - 60px)',
    overflow: 'auto',
    padding: '16px'
  });

  const eventCardStyles = mergeStyles({
    backgroundColor: theme.palette.white,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: '6px',
    padding: '12px',
    marginBottom: '8px',
    cursor: 'pointer',
    transition: 'all 0.2s ease',
    '&:hover': {
      backgroundColor: theme.palette.neutralLighterAlt,
      borderColor: theme.palette.themePrimary,
      boxShadow: theme.effects.elevation4
    }
  });

  const groupHeaderStyles = mergeStyles({
    backgroundColor: theme.palette.themeLighter,
    padding: '8px 12px',
    borderRadius: '4px',
    marginBottom: '8px',
    marginTop: '16px',
    '&:first-child': {
      marginTop: '0'
    }
  });

  const groupByOptions: IDropdownOption[] = [
    { key: 'date', text: 'Group by Date' },
    { key: 'calendar', text: 'Group by Calendar' },
    { key: 'category', text: 'Group by Category' }
  ];

  const daysOptions: IDropdownOption[] = [
    { key: 7, text: 'Next 7 days' },
    { key: 14, text: 'Next 14 days' },
    { key: 30, text: 'Next 30 days' },
    { key: 0, text: 'All events' }
  ];

  // Filter events based on selected days
  const filteredEvents = useMemo(() => {
    if (selectedDays === 0) return events;
    
    const now = new Date();
    const futureDate = moment(now).add(selectedDays, 'days').toDate();
    
    return events.filter(event => 
      event.start >= now && event.start <= futureDate
    );
  }, [events, selectedDays]);

  // Group events based on selected grouping
  const groupedEvents = useMemo(() => {
    const groups: { [key: string]: ICalendarEvent[] } = {};

    filteredEvents.forEach(event => {
      let groupKey: string;

      switch (selectedGroupBy) {
        case 'calendar': {
          groupKey = event.calendarTitle;
          break;
        }
        case 'category': {
          groupKey = event.category || 'Uncategorized';
          break;
        }
        case 'date':
        default: {
          groupKey = DateUtils.formatDate(event.start, 'YYYY-MM-DD');
          break;
        }
      }

      if (!groups[groupKey]) {
        groups[groupKey] = [];
      }
      groups[groupKey].push(event);
    });

    // Sort groups and events within groups
    const sortedGroups: { [key: string]: ICalendarEvent[] } = {};
    const sortedKeys = Object.keys(groups).sort((a, b) => {
      if (selectedGroupBy === 'date') {
        return new Date(a).getTime() - new Date(b).getTime();
      }
      return a.localeCompare(b);
    });

    sortedKeys.forEach(key => {
      sortedGroups[key] = groups[key].sort((a, b) => 
        a.start.getTime() - b.start.getTime()
      );
    });

    return sortedGroups;
  }, [filteredEvents, selectedGroupBy]);

  const getGroupDisplayName = (groupKey: string): string => {
    switch (selectedGroupBy) {
      case 'date': {
        const date = new Date(groupKey);
        if (DateUtils.isToday(date)) return 'Today';
        if (DateUtils.isTomorrow(date)) return 'Tomorrow';
        return DateUtils.getFriendlyDate(date);
      }
      case 'calendar':
      case 'category':
      default: {
        return groupKey;
      }
    }
  };

  const getEventTimeDisplay = (event: ICalendarEvent): string => {
    if (event.isAllDay) {
      return 'All day';
    }
    
    if (selectedGroupBy === 'date') {
      return `${DateUtils.formatTime(event.start)} - ${DateUtils.formatTime(event.end)}`;
    }
    
    return `${DateUtils.formatDateTime(event.start)} - ${DateUtils.formatTime(event.end)}`;
  };

  const renderEventCard = (event: ICalendarEvent): React.ReactElement => {
    const calendarSource = calendarSources.find(s => s.id === event.calendarId);
    const eventColor = event.color || calendarSource?.color || theme.palette.themePrimary;

    return (
      <div
        key={event.id}
        className={eventCardStyles}
        onClick={() => onEventSelect(event)}
        style={{ borderLeft: `4px solid ${eventColor}` }}
      >
        <Stack tokens={{ childrenGap: 6 }}>
          {/* Event Title and Time */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
            <Stack grow tokens={{ childrenGap: 4 }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                {event.title}
              </Text>
              
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon
                  iconName={event.isAllDay ? 'DateTime' : 'Clock'}
                  styles={{
                    root: {
                      fontSize: '12px',
                      color: theme.palette.neutralSecondary
                    }
                  }}
                />
                <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  {getEventTimeDisplay(event)}
                </Text>
              </Stack>
            </Stack>

            {/* Event Status Indicators */}
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              {event.importance === 'high' && (
                <Icon
                  iconName="Important"
                  styles={{
                    root: {
                      fontSize: '14px',
                      color: theme.palette.red
                    }
                  }}
                  title="High Priority"
                />
              )}
              
              {event.isRecurring && (
                <Icon
                  iconName="Sync"
                  styles={{
                    root: {
                      fontSize: '12px',
                      color: theme.palette.neutralSecondary
                    }
                  }}
                  title="Recurring Event"
                />
              )}

              {event.attendees && event.attendees.length > 0 && (
                <Icon
                  iconName="People"
                  styles={{
                    root: {
                      fontSize: '12px',
                      color: theme.palette.neutralSecondary
                    }
                  }}
                  title={`${event.attendees.length} attendee${event.attendees.length !== 1 ? 's' : ''}`}
                />
              )}
            </Stack>
          </Stack>

          {/* Location */}
          {event.location && (
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
              <Icon
                iconName="MapPin"
                styles={{
                  root: {
                    fontSize: '12px',
                    color: theme.palette.neutralSecondary
                  }
                }}
              />
              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {event.location}
              </Text>
            </Stack>
          )}

          {/* Description preview */}
          {event.description && (
            <Text
              variant="small"
              styles={{
                root: {
                  color: theme.palette.neutralSecondary,
                  display: '-webkit-box',
                  WebkitLineClamp: 2,
                  WebkitBoxOrient: 'vertical',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis'
                }
              }}
            >
              {event.description}
            </Text>
          )}

          {/* Calendar and Category info */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
              <Icon
                iconName={event.calendarType === 'SharePoint' ? 'SharePointLogo' : 'OutlookLogo'}
                styles={{
                  root: {
                    fontSize: '12px',
                    color: eventColor
                  }
                }}
              />
              <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {selectedGroupBy !== 'calendar' ? event.calendarTitle : ''}
                {selectedGroupBy !== 'category' && event.category ? ` • ${event.category}` : ''}
              </Text>
            </Stack>

            {event.webUrl && (
              <Link
                href={event.webUrl}
                target="_blank"
                styles={{ root: { fontSize: '11px' } }}
                onClick={(e) => e.stopPropagation()}
              >
                <Icon iconName="OpenInNewWindow" />
              </Link>
            )}
          </Stack>
        </Stack>
      </div>
    );
  };

  return (
    <div className={containerStyles}>
      {/* Header */}
      <div className={headerStyles}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            Agenda View
          </Text>
          
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <Text variant="medium" styles={{ root: { color: theme.palette.neutralSecondary } }}>
              {filteredEvents.length} events
            </Text>
            
            <Dropdown
              options={daysOptions}
              selectedKey={selectedDays}
              onChange={(_, option) => option && setSelectedDays(option.key as number)}
              styles={{ root: { minWidth: '120px' } }}
            />
            
            <Dropdown
              options={groupByOptions}
              selectedKey={selectedGroupBy}
              onChange={(_, option) => option && setSelectedGroupBy(option.key as string)}
              styles={{ root: { minWidth: '150px' } }}
            />
          </Stack>
        </Stack>
      </div>

      {/* Content */}
      <div className={contentStyles}>
        {Object.keys(groupedEvents).length === 0 ? (
          <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '300px' } }}>
            <Icon 
              iconName="Calendar" 
              styles={{ 
                root: { 
                  fontSize: '48px', 
                  color: theme.palette.neutralTertiary, 
                  marginBottom: '16px' 
                } 
              }} 
            />
            <Text variant="large" styles={{ root: { color: theme.palette.neutralSecondary } }}>
              No events found
            </Text>
            <Text variant="medium" styles={{ root: { color: theme.palette.neutralTertiary } }}>
              Try adjusting your date range or check your calendar sources
            </Text>
          </Stack>
        ) : (
          <Stack tokens={stackTokens}>
            {Object.entries(groupedEvents).map(([groupKey, groupEvents]) => (
              <Stack key={groupKey} tokens={{ childrenGap: 8 }}>
                {/* Group Header */}
                <div className={groupHeaderStyles}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="medium" styles={{ root: { fontWeight: 600, color: theme.palette.themePrimary } }}>
                      {getGroupDisplayName(groupKey)}
                    </Text>
                    <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                      {groupEvents.length} event{groupEvents.length !== 1 ? 's' : ''}
                    </Text>
                  </Stack>
                </div>

                {/* Group Events */}
                <Stack tokens={{ childrenGap: 4 }}>
                  {groupEvents.map(event => renderEventCard(event))}
                </Stack>
              </Stack>
            ))}
          </Stack>
        )}
      </div>
    </div>
  );
};