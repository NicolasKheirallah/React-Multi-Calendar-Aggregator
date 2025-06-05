import * as React from 'react';
import { useState, useEffect } from 'react';
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
  ITheme,
  ProgressIndicator,
  Toggle,
  CommandBar,
  ICommandBarItemProps,
  TooltipHost,
  DirectionalHint
} from '@fluentui/react';

import { ICalendarSource, CalendarSourceType } from '../models/ICalendarModels';
import { DateUtils } from '../utils/DateUtils';

export interface ICalendarSourcesPanelProps {
  sources: ICalendarSource[];
  selectedSources: string[];
  onSourcesChange: (selectedIds: string[]) => void;
  onClose: () => void;
  onRefresh?: () => void;
}

interface ISourceWithHealth extends ICalendarSource {
  isHealthy?: boolean;
  healthStatus?: 'checking' | 'healthy' | 'warning' | 'error';
  healthMessage?: string;
  responseTime?: number;
  lastChecked?: Date;
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
  const [filterHealth, setFilterHealth] = useState<string>('all');
  const [saving, setSaving] = useState<boolean>(false);
  const [sourcesWithHealth, setSourcesWithHealth] = useState<ISourceWithHealth[]>([]);
  const [isCheckingHealth, setIsCheckingHealth] = useState<boolean>(false);
  const [showAdvanced, setShowAdvanced] = useState<boolean>(false);
  const [sortColumn, setSortColumn] = useState<string>('title');
  const [sortDescending, setSortDescending] = useState<boolean>(false);

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

  const healthIndicatorStyles = mergeStyles({
    width: '12px',
    height: '12px',
    borderRadius: '50%',
    display: 'inline-block',
    marginRight: '8px'
  });

  // Initialize sources with health status
  useEffect(() => {
    setSourcesWithHealth(sources.map(source => ({
      ...source,
      healthStatus: 'checking' as const,
      healthMessage: 'Checking connection...'
    })));
  }, [sources]);

  // Check health of all sources
  const checkSourcesHealth = async (): Promise<void> => {
    setIsCheckingHealth(true);
    
    const updatedSources = await Promise.all(
      sourcesWithHealth.map(async (source) => {
        const startTime = Date.now();
        
        try {
          // Simulate health check with timeout
          const healthCheck = new Promise<boolean>((resolve) => {
            // Mock health check - in real implementation, this would ping the actual service
            setTimeout(() => {
              resolve(Math.random() > 0.1); // 90% success rate for demo
            }, Math.random() * 2000 + 500); // Random delay 500-2500ms
          });
          
          const timeoutPromise = new Promise<boolean>((_, reject) => {
            setTimeout(() => reject(new Error('Health check timeout')), 5000);
          });
          
          const isHealthy = await Promise.race([healthCheck, timeoutPromise]);
          const responseTime = Date.now() - startTime;
          
          return {
            ...source,
            isHealthy,
            healthStatus: (isHealthy ? 'healthy' : 'warning') as 'healthy' | 'warning',
            healthMessage: isHealthy ? `Connected (${responseTime}ms)` : 'Connection issues detected',
            responseTime,
            lastChecked: new Date()
          };
        } catch {
          const responseTime = Date.now() - startTime;
          return {
            ...source,
            isHealthy: false,
            healthStatus: 'error' as const,
            healthMessage: 'Connection failed',
            responseTime,
            lastChecked: new Date()
          };
        }
      })
    );
    
    setSourcesWithHealth(updatedSources);
    setIsCheckingHealth(false);
  };

  // Filter sources based on search, type, and health
  const filteredSources = React.useMemo(() => {
    let filtered = sourcesWithHealth;

    // Filter by search query
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      filtered = filtered.filter(source =>
        source.title.toLowerCase().includes(query) ||
        source.siteTitle.toLowerCase().includes(query) ||
        (source.description && source.description.toLowerCase().includes(query))
      );
    }

    // Filter by type
    if (filterType !== 'all') {
      filtered = filtered.filter(source => 
        source.type.toLowerCase() === filterType.toLowerCase()
      );
    }

    // Filter by health
    if (filterHealth !== 'all') {
      filtered = filtered.filter(source => 
        source.healthStatus === filterHealth
      );
    }

    // Sort sources
    filtered.sort((a, b) => {
      let comparison = 0;
      
      switch (sortColumn) {
        case 'title': {
          comparison = a.title.localeCompare(b.title);
          break;
        }
        case 'type': {
          comparison = a.type.localeCompare(b.type);
          break;
        }
        case 'health': {
          const healthOrder = { healthy: 0, warning: 1, error: 2, checking: 3 };
          comparison = (healthOrder[a.healthStatus || 'checking'] || 3) - (healthOrder[b.healthStatus || 'checking'] || 3);
          break;
        }
        case 'responseTime': {
          comparison = (a.responseTime || 9999) - (b.responseTime || 9999);
          break;
        }
        default: {
          comparison = a.title.localeCompare(b.title);
        }
      }
      
      return sortDescending ? -comparison : comparison;
    });

    return filtered;
  }, [sourcesWithHealth, searchQuery, filterType, filterHealth, sortColumn, sortDescending]);

  const typeFilterOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Types' },
    { key: 'sharepoint', text: 'SharePoint Calendars' },
    { key: 'exchange', text: 'Exchange Calendars' }
  ];

  const healthFilterOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Status' },
    { key: 'healthy', text: 'Healthy' },
    { key: 'warning', text: 'Warning' },
    { key: 'error', text: 'Error' },
    { key: 'checking', text: 'Checking' }
  ];

  const sortOptions: IDropdownOption[] = [
    { key: 'title', text: 'Name' },
    { key: 'type', text: 'Type' },
    { key: 'health', text: 'Health' },
    { key: 'responseTime', text: 'Response Time' }
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

  const handleSelectHealthy = (): void => {
    const healthyIds = filteredSources
      .filter(s => s.healthStatus === 'healthy')
      .map(s => s.id);
    setLocalSelectedSources([...new Set([...localSelectedSources, ...healthyIds])]);
  };

  const handleSave = async (): Promise<void> => {
    setSaving(true);
    try {
      onSourcesChange(localSelectedSources);
      // Add small delay to show saving state
      await new Promise(resolve => setTimeout(resolve, 500));
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

  const getHealthColor = (status?: string): string => {
    switch (status) {
      case 'healthy':
        return theme.palette.green;
      case 'warning':
        return theme.palette.yellow;
      case 'error':
        return theme.palette.red;
      case 'checking':
      default:
        return theme.palette.neutralSecondary;
    }
  };

  const getSourceStats = (): {
    total: number;
    sharePoint: number;
    exchange: number;
    selected: number;
    healthy: number;
    warning: number;
    error: number;
  } => {
    const sharePointCount = sourcesWithHealth.filter(s => s.type === CalendarSourceType.SharePoint).length;
    const exchangeCount = sourcesWithHealth.filter(s => s.type === CalendarSourceType.Exchange).length;
    const selectedCount = localSelectedSources.length;
    const healthyCount = sourcesWithHealth.filter(s => s.healthStatus === 'healthy').length;
    const warningCount = sourcesWithHealth.filter(s => s.healthStatus === 'warning').length;
    const errorCount = sourcesWithHealth.filter(s => s.healthStatus === 'error').length;

    return {
      total: sourcesWithHealth.length,
      sharePoint: sharePointCount,
      exchange: exchangeCount,
      selected: selectedCount,
      healthy: healthyCount,
      warning: warningCount,
      error: errorCount
    };
  };

  const stats = getSourceStats();

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'selectAll',
      text: 'Select All',
      iconProps: { iconName: 'CheckboxComposite' },
      onClick: handleSelectAll
    },
    {
      key: 'deselectAll',
      text: 'Deselect All',
      iconProps: { iconName: 'Checkbox' },
      onClick: handleDeselectAll
    },
    {
      key: 'selectHealthy',
      text: 'Select Healthy',
      iconProps: { iconName: 'Health' },
      onClick: handleSelectHealthy
    },
    {
      key: 'checkHealth',
      text: 'Check Health',
      iconProps: { iconName: 'Heart' },
      disabled: isCheckingHealth,
      onClick: (): void => {
        checkSourcesHealth().catch(console.error);
      }
    }
  ];

  const renderSourceCard = (source: ISourceWithHealth): React.ReactElement => {
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
            
            <Stack grow tokens={{ childrenGap: 8 }}>
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
                
                {/* Health Indicator */}
                <TooltipHost
                  content={source.healthMessage || 'Unknown status'}
                  directionalHint={DirectionalHint.topCenter}
                >
                  <div
                    className={healthIndicatorStyles}
                    style={{ backgroundColor: getHealthColor(source.healthStatus) }}
                  />
                </TooltipHost>
                
                {source.healthStatus === 'checking' && (
                  <Spinner size={SpinnerSize.xSmall} />
                )}
              </Stack>

              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {source.description}
              </Text>

              <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
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

                {source.responseTime && (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                    <Icon
                      iconName="Speedometer"
                      styles={{ root: { fontSize: '12px', color: theme.palette.neutralSecondary } }}
                    />
                    <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                      {source.responseTime}ms
                    </Text>
                  </Stack>
                )}
              </Stack>
              
              {source.lastChecked && (
                <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary, fontStyle: 'italic' } }}>
                  Last checked: {DateUtils.getRelativeTime(source.lastChecked)}
                </Text>
              )}
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

      {/* Health Check Progress */}
      {isCheckingHealth && (
        <ProgressIndicator 
          description="Checking calendar source health..."
          styles={{ root: { marginBottom: '16px' } }}
        />
      )}

      {/* Statistics */}
      <MessageBar messageBarType={MessageBarType.info}>
        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          <Text variant="small">
            <strong>{stats.total}</strong> sources found
          </Text>
          <Text variant="small">
            <strong>{stats.sharePoint}</strong> SharePoint
          </Text>
          <Text variant="small">
            <strong>{stats.exchange}</strong> Exchange
          </Text>
          <Text variant="small">
            <strong>{stats.selected}</strong> selected
          </Text>
          <Text variant="small" styles={{ root: { color: theme.palette.green } }}>
            <strong>{stats.healthy}</strong> healthy
          </Text>
          {stats.warning > 0 && (
            <Text variant="small" styles={{ root: { color: theme.palette.yellow } }}>
              <strong>{stats.warning}</strong> warnings
            </Text>
          )}
          {stats.error > 0 && (
            <Text variant="small" styles={{ root: { color: theme.palette.red } }}>
              <strong>{stats.error}</strong> errors
            </Text>
          )}
        </Stack>
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

          <Dropdown
            placeholder="Filter by health..."
            options={healthFilterOptions}
            selectedKey={filterHealth}
            onChange={(_, option) => option && setFilterHealth(option.key as string)}
            styles={{ root: { minWidth: '150px' } }}
          />

          <Toggle
            label="Advanced view"
            checked={showAdvanced}
            onChange={(_, checked) => setShowAdvanced(checked || false)}
          />

          {onRefresh && (
            <DefaultButton
              iconProps={{ iconName: 'Refresh' }}
              text="Refresh Sources"
              onClick={onRefresh}
            />
          )}
        </Stack>

        {/* Advanced Controls */}
        {showAdvanced && (
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            <Dropdown
              label="Sort by"
              options={sortOptions}
              selectedKey={sortColumn}
              onChange={(_, option) => option && setSortColumn(option.key as string)}
              styles={{ root: { minWidth: '120px' } }}
            />
            
            <Toggle
              label="Descending"
              checked={sortDescending}
              onChange={(_, checked) => setSortDescending(checked || false)}
            />
          </Stack>
        )}

        {/* Bulk Actions */}
        {filteredSources.length > 0 && (
          <CommandBar
            items={commandBarItems}
            styles={{
              root: {
                backgroundColor: theme.palette.neutralLighterAlt,
                borderRadius: '4px',
                padding: '4px'
              }
            }}
          />
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
              {searchQuery || filterType !== 'all' || filterHealth !== 'all' ? 'No calendars match your filters' : 'No calendar sources found'}
            </Text>
            <Text variant="medium" styles={{ root: { color: theme.palette.neutralTertiary } }}>
              {searchQuery || filterType !== 'all' || filterHealth !== 'all' ? 'Try adjusting your search or filters' : 'Check your permissions and try refreshing'}
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
              onClick={(): void => {
                handleSave().catch(console.error);
              }}
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