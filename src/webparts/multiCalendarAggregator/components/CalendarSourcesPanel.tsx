import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  Checkbox,
  PrimaryButton,
  DefaultButton,
  Separator,
  Icon,
  Link,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  mergeStyles,
  useTheme,
  IStackTokens,
  SearchBox,
  Dropdown,
  IDropdownOption,
  ITheme
} from '@fluentui/react';

import { ICalendarSource, CalendarSourceType } from '../models/ICalendarModels';

export interface ICalendarSourcesPanelProps {
  sources: ICalendarSource[];
  selectedSources: string[];
  onSourcesChange: (selectedIds: string[]) => void;
  onClose: () => void;
  onRefresh?: () => void;
}

const stackTokens: IStackTokens = { childrenGap: 16 };

export const CalendarSourcesPanel: React.FC<ICalendarSourcesPanelProps> = ({
  sources,
  selectedSources,
  onSourcesChange,
  onClose,
  onRefresh
}) => {
  const theme: ITheme = useTheme();
  const [localSelectedSources, setLocalSelectedSources] = useState<string[]>(selectedSources);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterType, setFilterType] = useState<string>('all');
  const [saving, setSaving] = useState<boolean>(false);

  const containerStyles = mergeStyles({
    padding: '20px',
    height: '100%',
    overflow: 'auto'
  });

  const headerStyles = mergeStyles({
    backgroundColor: theme.palette.themePrimary,
    color: theme.palette.white,
    padding: '16px',
    margin: '-20px -20px 20px -20px',
    borderRadius: '0 0 8px 8px'
  });

  const sourceCardStyles = mergeStyles({
    backgroundColor: theme.palette.white,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: '6px',
    padding: '12px',
    marginBottom: '8px',
    transition: 'all 0.2s ease',
    '&:hover': {
      backgroundColor: theme.palette.neutralLighterAlt,
      borderColor: theme.palette.themePrimary,
      boxShadow: theme.effects.elevation4
    }
  });

  // Filter sources based on search and type
  const filteredSources = React.useMemo(() => {
    let filtered = sources;

    // Filter by search query
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      filtered = sources.filter(source =>
        source.title.toLowerCase().includes(query) ||
        source.siteTitle.toLowerCase().includes(query) ||
        source.description.toLowerCase().includes(query)
      );
    }

    // Filter by type
    if (filterType !== 'all') {
      filtered = filtered.filter(source => 
        source.type.toLowerCase() === filterType.toLowerCase()
      );
    }

    return filtered;
  }, [sources, searchQuery, filterType]);

  const typeFilterOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Types' },
    { key: 'sharepoint', text: 'SharePoint Calendars' },
    { key: 'exchange', text: 'Exchange Calendars' }
  ];

  const handleSourceToggle = (sourceId: string, checked: boolean): void => {
    const updated = checked 
      ? [...localSelectedSources, sourceId]
      : localSelectedSources.filter(id => id !== sourceId);
    setLocalSelectedSources(updated);
  };

  const handleSelectAll = (): void => {
    const allFilteredIds = filteredSources.map(s => s.id);
    setLocalSelectedSources([...new Set([...localSelectedSources, ...allFilteredIds])]);
  };

  const handleDeselectAll = (): void => {
    const filteredIds = new Set(filteredSources.map(s => s.id));
    setLocalSelectedSources(localSelectedSources.filter(id => !filteredIds.has(id)));
  };

  const handleSave = async (): Promise<void> => {
    setSaving(true);
    try {
      onSourcesChange(localSelectedSources);
      onClose();
    } catch (error) {
      console.error('Error saving calendar sources:', error);
    } finally {
      setSaving(false);
    }
  };

  const getSourceIcon = (type: CalendarSourceType): string => {
    switch (type) {
      case CalendarSourceType.Exchange:
        return 'OutlookLogo';
      case CalendarSourceType.SharePoint:
      default:
        return 'SharePointLogo';
    }
  };

  const getSourceStats = (): {
    total: number;
    sharePoint: number;
    exchange: number;
    selected: number;
  } => {
    const sharePointCount = sources.filter(s => s.type === CalendarSourceType.SharePoint).length;
    const exchangeCount = sources.filter(s => s.type === CalendarSourceType.Exchange).length;
    const selectedCount = localSelectedSources.length;

    return {
      total: sources.length,
      sharePoint: sharePointCount,
      exchange: exchangeCount,
      selected: selectedCount
    };
  };

  const stats = getSourceStats();

  const renderSourceCard = (source: ICalendarSource): React.ReactElement => {
    const isSelected = localSelectedSources.includes(source.id);
    
    return (
      <div key={source.id} className={sourceCardStyles}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
          <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 12 }} grow>
            <Checkbox
              checked={isSelected}
              onChange={(_, checked) => handleSourceToggle(source.id, checked || false)}
              styles={{ root: { marginTop: '2px' } }}
            />
            
            <Stack grow tokens={{ childrenGap: 6 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon
                  iconName={getSourceIcon(source.type)}
                  styles={{
                    root: {
                      fontSize: '16px',
                      color: source.color || theme.palette.themePrimary
                    }
                  }}
                />
                <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                  {source.title}
                </Text>
              </Stack>

              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {source.description}
              </Text>

              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Icon
                    iconName="WebAppBuilderFragment"
                    styles={{ root: { fontSize: '12px', color: theme.palette.neutralSecondary } }}
                  />
                  <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    {source.siteTitle}
                  </Text>
                </Stack>

                {source.itemCount !== undefined && (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                    <Icon
                      iconName="NumberField"
                      styles={{ root: { fontSize: '12px', color: theme.palette.neutralSecondary } }}
                    />
                    <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                      {source.itemCount} events
                    </Text>
                  </Stack>
                )}

                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                  <Icon
                    iconName="Shield"
                    styles={{ root: { fontSize: '12px', color: theme.palette.neutralSecondary } }}
                  />
                  <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    {source.type}
                  </Text>
                </Stack>
              </Stack>
            </Stack>
          </Stack>

          <Stack horizontalAlign="end" tokens={{ childrenGap: 4 }}>
            <Link
              href={source.url}
              target="_blank"
              styles={{ root: { fontSize: '12px' } }}
            >
              <Icon iconName="OpenInNewWindow" />
            </Link>
            
            {source.lastModified && (
              <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary } }}>
                {new Date(source.lastModified).toLocaleDateString()}
              </Text>
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
          <Stack>
            <Text variant="xLarge" styles={{ root: { fontWeight: 600, color: theme.palette.white } }}>
              Manage Calendar Sources
            </Text>
            <Text variant="medium" styles={{ root: { color: theme.palette.white, opacity: 0.9 } }}>
              Select calendars to display in the aggregated view
            </Text>
          </Stack>
          <Icon
            iconName="Calendar"
            styles={{
              root: {
                fontSize: '24px',
                color: theme.palette.white,
                opacity: 0.8
              }
            }}
          />
        </Stack>
      </div>

      {/* Statistics */}
      <MessageBar messageBarType={MessageBarType.info}>
        <Text variant="small">
          <strong>{stats.total}</strong> calendar sources found: 
          <strong> {stats.sharePoint}</strong> SharePoint, 
          <strong> {stats.exchange}</strong> Exchange • 
          <strong> {stats.selected}</strong> selected
        </Text>
      </MessageBar>

      {/* Filters and Controls */}
      <Stack tokens={stackTokens}>
        <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
          <SearchBox
            placeholder="Search calendar sources..."
            value={searchQuery}
            onChange={(_, value) => setSearchQuery(value || '')}
            styles={{ root: { minWidth: '200px', maxWidth: '300px' } }}
          />
          
          <Dropdown
            placeholder="Filter by type..."
            options={typeFilterOptions}
            selectedKey={filterType}
            onChange={(_, option) => option && setFilterType(option.key as string)}
            styles={{ root: { minWidth: '150px' } }}
          />

          {onRefresh && (
            <DefaultButton
              iconProps={{ iconName: 'Refresh' }}
              text="Refresh"
              onClick={onRefresh}
            />
          )}
        </Stack>

        {/* Bulk Actions */}
        {filteredSources.length > 0 && (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              iconProps={{ iconName: 'CheckboxComposite' }}
              text="Select All Filtered"
              onClick={handleSelectAll}
            />
            <DefaultButton
              iconProps={{ iconName: 'Checkbox' }}
              text="Deselect All Filtered"
              onClick={handleDeselectAll}
            />
          </Stack>
        )}

        <Separator />

        {/* Calendar Sources List */}
        {filteredSources.length === 0 ? (
          <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '200px' } }}>
            <Icon 
              iconName="SearchAndApps" 
              styles={{ root: { fontSize: '48px', color: theme.palette.neutralTertiary, marginBottom: '16px' } }} 
            />
            <Text variant="large" styles={{ root: { color: theme.palette.neutralSecondary } }}>
              {searchQuery || filterType !== 'all' ? 'No calendars match your filters' : 'No calendar sources found'}
            </Text>
            <Text variant="medium" styles={{ root: { color: theme.palette.neutralTertiary } }}>
              {searchQuery || filterType !== 'all' ? 'Try adjusting your search or filters' : 'Check your permissions and try refreshing'}
            </Text>
          </Stack>
        ) : (
          <Stack tokens={{ childrenGap: 8 }}>
            {filteredSources.map(source => renderSourceCard(source))}
          </Stack>
        )}

        <Separator />

        {/* Action Buttons */}
        <Stack horizontal horizontalAlign="space-between">
          <DefaultButton
            text="Cancel"
            onClick={onClose}
          />
          
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <Text variant="medium" styles={{ root: { alignSelf: 'center', color: theme.palette.neutralSecondary } }}>
              {localSelectedSources.length} source{localSelectedSources.length !== 1 ? 's' : ''} selected
            </Text>
            
            <PrimaryButton
              text={saving ? 'Saving...' : 'Save Changes'}
              onClick={handleSave}
              disabled={saving}
              iconProps={saving ? undefined : { iconName: 'Save' }}
            />
            
            {saving && <Spinner size={SpinnerSize.small} />}
          </Stack>
        </Stack>
      </Stack>
    </div>
  );
};