import * as React from 'react';
import {
  Stack,
  Text,
  Icon,
  Link,
  Separator,
  mergeStyles,
  useTheme,
  IStackTokens,
  Label,
  ITheme
} from '@fluentui/react';
import { DateUtils } from '../utils/DateUtils';
import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';

export interface IEventDetailsPanelProps {
  event: ICalendarEvent;
  calendarSource?: ICalendarSource;
  onClose: () => void;
  onEdit?: (event: ICalendarEvent) => void;
  onDelete?: (eventId: string) => void;
}

const stackTokens: IStackTokens = { childrenGap: 12 };

export const EventDetailsPanel: React.FC<IEventDetailsPanelProps> = ({
  event,
  calendarSource,
  onClose,
  onEdit,
  onDelete
}) => {
  const theme: ITheme = useTheme();

  const containerStyles = mergeStyles({
    padding: '20px',
    height: '100%',
    overflow: 'auto'
  });

  const headerStyles = mergeStyles({
    backgroundColor: event.color || theme.palette.themePrimary,
    color: theme.palette.white,
    padding: '16px',
    margin: '-20px -20px 20px -20px',
    borderRadius: '0 0 8px 8px'
  });

  const sectionStyles = mergeStyles({
    backgroundColor: theme.palette.neutralLighterAlt,
    padding: '12px 16px',
    borderRadius: '6px',
    border: `1px solid ${theme.palette.neutralLight}`
  });

  const iconStyles = {
    root: {
      fontSize: '16px',
      color: theme.palette.themePrimary,
      marginRight: '8px'
    }
  };

  const formatDateTime = (date: Date): string => {
    if (event.isAllDay) {
      return DateUtils.formatDate(date);
    }
    return DateUtils.formatDateTime(date);
  };

  const getDuration = (): string => {
    if (event.isAllDay) {
      const days = Math.ceil((event.end.getTime() - event.start.getTime()) / (1000 * 60 * 60 * 24));
      return days === 1 ? 'All day' : `${days} days`;
    }
    return DateUtils.getDuration(event.start, event.end);
  };

  return (
    <div className={containerStyles}>
      {/* Header */}
      <div className={headerStyles}>
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="xLarge" styles={{ root: { fontWeight: 600, color: theme.palette.white } }}>
            {event.title}
          </Text>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName="Calendar" styles={{ root: { color: theme.palette.white } }} />
            <Text variant="medium" styles={{ root: { color: theme.palette.white, opacity: 0.9 } }}>
              {calendarSource?.title || event.calendarTitle}
            </Text>
          </Stack>
        </Stack>
      </div>

      <Stack tokens={stackTokens}>
        {/* Date and Time */}
        <div className={sectionStyles}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>Date & Time</Label>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="DateTime" styles={iconStyles} />
              <Stack>
                <Text variant="medium" styles={{ root: { fontWeight: 500 } }}>
                  {formatDateTime(event.start)}
                </Text>
                {!event.isAllDay && (
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    to {DateUtils.formatDateTime(event.end)}
                  </Text>
                )}
                <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  Duration: {getDuration()}
                </Text>
              </Stack>
            </Stack>
          </Stack>
        </div>

        {/* Description */}
        {event.description && (
          <div className={sectionStyles}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Label>Description</Label>
              <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }}>
                <Icon iconName="FileComment" styles={iconStyles} />
                <Text variant="medium" styles={{ root: { lineHeight: 1.5 } }}>
                  {event.description}
                </Text>
              </Stack>
            </Stack>
          </div>
        )}

        {/* Location */}
        {event.location && (
          <div className={sectionStyles}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Label>Location</Label>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="MapPin" styles={iconStyles} />
                <Text variant="medium">{event.location}</Text>
              </Stack>
            </Stack>
          </div>
        )}

        {/* Organizer */}
        <div className={sectionStyles}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>Organizer</Label>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="Contact" styles={iconStyles} />
              <Text variant="medium">{event.organizer}</Text>
            </Stack>
          </Stack>
        </div>

        {/* Attendees */}
        {event.attendees && event.attendees.length > 0 && (
          <div className={sectionStyles}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Label>Attendees ({event.attendees.length})</Label>
              <Stack tokens={{ childrenGap: 4 }}>
                {event.attendees.slice(0, 5).map((attendee, index) => (
                  <Stack key={index} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Icon iconName="Contact" styles={iconStyles} />
                    <Stack>
                      <Text variant="medium">{attendee.name}</Text>
                      <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                        {attendee.email} • {attendee.response}
                      </Text>
                    </Stack>
                  </Stack>
                ))}
                {event.attendees.length > 5 && (
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, marginLeft: '24px' } }}>
                    ... and {event.attendees.length - 5} more
                  </Text>
                )}
              </Stack>
            </Stack>
          </div>
        )}

        {/* Properties */}
        <div className={sectionStyles}>
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>Properties</Label>
            
            <Stack tokens={{ childrenGap: 8 }}>
              {/* Category */}
              {event.category && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName="Tag" styles={iconStyles} />
                  <Text variant="medium">Category: {event.category}</Text>
                </Stack>
              )}

              {/* Importance */}
              {event.importance && event.importance !== 'normal' && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon 
                    iconName={event.importance === 'high' ? 'Important' : 'StatusCircleRing'} 
                    styles={{
                      root: {
                        fontSize: '16px',
                        color: event.importance === 'high' ? theme.palette.red : theme.palette.neutralSecondary,
                        marginRight: '8px'
                      }
                    }} 
                  />
                  <Text variant="medium">
                    Priority: {event.importance.charAt(0).toUpperCase() + event.importance.slice(1)}
                  </Text>
                </Stack>
              )}

              {/* All Day */}
              {event.isAllDay && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName="DateTime" styles={iconStyles} />
                  <Text variant="medium">All day event</Text>
                </Stack>
              )}

              {/* Recurring */}
              {event.isRecurring && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName="Sync" styles={iconStyles} />
                  <Text variant="medium">Recurring event</Text>
                </Stack>
              )}

              {/* Calendar Source */}
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon 
                  iconName={event.calendarType === 'SharePoint' ? 'SharePointLogo' : 'OutlookLogo'} 
                  styles={iconStyles} 
                />
                <Text variant="medium">Source: {event.calendarType}</Text>
              </Stack>
            </Stack>
          </Stack>
        </div>

        {/* Timestamps */}
        <div className={sectionStyles}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>Event Details</Label>
            <Stack tokens={{ childrenGap: 4 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Add" styles={iconStyles} />
                <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  Created: {DateUtils.formatDateTime(event.created)}
                </Text>
              </Stack>
              
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Edit" styles={iconStyles} />
                <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  Modified: {DateUtils.formatDateTime(event.modified)}
                </Text>
              </Stack>

              {event.webUrl && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName="Link" styles={iconStyles} />
                  <Link href={event.webUrl} target="_blank" styles={{ root: { fontSize: '12px' } }}>
                    View in {event.calendarType}
                  </Link>
                </Stack>
              )}
            </Stack>
          </Stack>
        </div>

        <Separator />

        {/* Actions */}
        {(onEdit || onDelete) && (
          <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center">
            {onEdit && (
              <Link onClick={() => onEdit(event)} styles={{ root: { fontSize: '14px' } }}>
                <Icon iconName="Edit" styles={{ root: { marginRight: '4px' } }} />
                Edit Event
              </Link>
            )}
            
            {onDelete && (
              <Link 
                onClick={() => onDelete(event.id)} 
                styles={{ 
                  root: { 
                    fontSize: '14px',
                    color: theme.palette.red 
                  } 
                }}
              >
                <Icon iconName="Delete" styles={{ root: { marginRight: '4px', color: theme.palette.red } }} />
                Delete Event
              </Link>
            )}
          </Stack>
        )}
      </Stack>
    </div>
  );
};