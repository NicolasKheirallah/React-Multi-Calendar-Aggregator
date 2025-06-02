import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneLink,
  PropertyPaneLabel,
  PropertyPaneButton,
  PropertyPaneButtonType,
  IPropertyPaneConfiguration,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'MultiCalendarAggregatorWebPartStrings';
import MultiCalendarAggregator from './components/MultiCalendarAggregator';
import { IMultiCalendarAggregatorProps } from './components/IMultiCalendarAggregatorProps';
import { CalendarViewType } from './models/ICalendarModels';
import { AppConstants } from './constants/AppConstants';

export interface IMultiCalendarAggregatorWebPartProps {
  // Basic Configuration
  title: string;
  description: string;
  selectedCalendars: string[];
  viewType: CalendarViewType;
  
  // Display Options
  showWeekends: boolean;
  maxEvents: number;
  colorCoding: boolean;
  showEventDetails: boolean;
  compactView: boolean;
  
  // Data Sources
  useGraphAPI: boolean;
  includePersonalCalendars: boolean;
  includeSharedCalendars: boolean;
  includeGroupCalendars: boolean;
  
  // Performance & Caching
  refreshInterval: number;
  enableCaching: boolean;
  cacheTimeout: number;
  
  // Advanced Features
  enableConflictDetection: boolean;
  enableSearch: boolean;
  enableFilters: boolean;
  enableExport: boolean;
  
  // Accessibility & UX
  enableKeyboardNavigation: boolean;
  enableHighContrast: boolean;
  showLoadingAnimation: boolean;
  
  // Developer Options
  enableDebugMode: boolean;
  logLevel: string;
  showPerformanceMetrics: boolean;
}

export default class MultiCalendarAggregatorWebPart extends BaseClientSideWebPart<IMultiCalendarAggregatorWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public override render(): void {
    // Validate properties before rendering
    this._validateProperties();

    const element: React.ReactElement<IMultiCalendarAggregatorProps> = React.createElement(
      MultiCalendarAggregator,
      {
        // Basic props
        title: this.properties.title || 'Multi-Calendar Aggregator',
        selectedCalendars: this.properties.selectedCalendars || [],
        viewType: this.properties.viewType as CalendarViewType || 'month',
        
        // Display options
        showWeekends: this.properties.showWeekends !== false, // Default to true
        maxEvents: this.properties.maxEvents || AppConstants.DEFAULT_MAX_EVENTS,
        colorCoding: this.properties.colorCoding !== false, // Default to true
        
        // Data source options
        useGraphAPI: this.properties.useGraphAPI || false,
        refreshInterval: this.properties.refreshInterval || AppConstants.DEFAULT_REFRESH_INTERVAL,
        
        // SharePoint Framework props
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected override onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            case 'Teams':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected override onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
  }

  protected override onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected override get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _validateProperties(): void {
    // Validate and fix property values
    if (!this.properties.title) {
      this.properties.title = 'Multi-Calendar Aggregator';
    }
    
    if (!this.properties.maxEvents || this.properties.maxEvents < AppConstants.VALIDATION.MIN_MAX_EVENTS) {
      this.properties.maxEvents = AppConstants.DEFAULT_MAX_EVENTS;
    }
    
    if (this.properties.maxEvents > AppConstants.VALIDATION.MAX_MAX_EVENTS) {
      this.properties.maxEvents = AppConstants.VALIDATION.MAX_MAX_EVENTS;
    }
    
    if (!this.properties.refreshInterval || this.properties.refreshInterval < 0) {
      this.properties.refreshInterval = AppConstants.DEFAULT_REFRESH_INTERVAL;
    }
    
    if (!this.properties.viewType) {
      this.properties.viewType = 'month' as CalendarViewType;
    }
    
    if (!this.properties.selectedCalendars) {
      this.properties.selectedCalendars = [];
    }
  }

  private _clearCache(): void {
    // Clear local storage cache
    const cacheKeys = Object.keys(localStorage).filter(key => 
      key.startsWith(AppConstants.CACHE_KEY_PREFIX)
    );
    cacheKeys.forEach(key => localStorage.removeItem(key));
    
    // Re-render to trigger data reload
    this.render();
  }

  private _exportConfiguration(): void {
    const config = {
      properties: this.properties,
      version: this.dataVersion.toString(),
      exportDate: new Date().toISOString(),
      webPartId: this.context.instanceId
    };
    
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `calendar-webpart-config-${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    URL.revokeObjectURL(url);
  }

  protected override getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            // Basic Configuration Group
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: this.properties.title,
                  maxLength: AppConstants.VALIDATION.MAX_TITLE_LENGTH,
                  placeholder: 'Enter web part title...'
                }),
                PropertyPaneTextField('description', {
                  label: 'Description',
                  value: this.properties.description,
                  multiline: true,
                  rows: 3,
                  placeholder: 'Optional description for the calendar aggregator...'
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneDropdown('viewType', {
                  label: 'Default View',
                  options: [
                    { key: 'month', text: 'Month View' },
                    { key: 'week', text: 'Week View' },
                    { key: 'day', text: 'Day View' },
                    { key: 'agenda', text: 'Agenda View' },
                    { key: 'timeline', text: 'Timeline View' }
                  ],
                  selectedKey: this.properties.viewType
                }),
                PropertyPaneSlider('maxEvents', {
                  label: 'Maximum Events to Display',
                  min: AppConstants.VALIDATION.MIN_MAX_EVENTS,
                  max: AppConstants.VALIDATION.MAX_MAX_EVENTS,
                  step: 10,
                  showValue: true,
                  value: this.properties.maxEvents
                })
              ]
            },
            
            // Display Options Group
            {
              groupName: 'Display Options',
              groupFields: [
                PropertyPaneToggle('showWeekends', {
                  label: 'Show Weekends',
                  checked: this.properties.showWeekends,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('colorCoding', {
                  label: 'Enable Color Coding',
                  checked: this.properties.colorCoding,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('showEventDetails', {
                  label: 'Show Event Details Panel',
                  checked: this.properties.showEventDetails !== false,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('compactView', {
                  label: 'Compact View',
                  checked: this.properties.compactView || false,
                  onText: 'Compact',
                  offText: 'Full'
                })
              ]
            }
          ]
        },
        
        // Data Sources Configuration Page
        {
          header: {
            description: 'Configure which calendar sources to include'
          },
          groups: [
            {
              groupName: 'Data Sources',
              groupFields: [
                PropertyPaneToggle('useGraphAPI', {
                  label: 'Microsoft Graph API (Exchange/Outlook)',
                  checked: this.properties.useGraphAPI,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('includePersonalCalendars', {
                  label: 'Include Personal Calendars',
                  checked: this.properties.includePersonalCalendars !== false,
                  onText: 'Include',
                  offText: 'Exclude',
                  disabled: !this.properties.useGraphAPI
                }),
                PropertyPaneToggle('includeSharedCalendars', {
                  label: 'Include Shared Calendars',
                  checked: this.properties.includeSharedCalendars !== false,
                  onText: 'Include',
                  offText: 'Exclude',
                  disabled: !this.properties.useGraphAPI
                }),
                PropertyPaneToggle('includeGroupCalendars', {
                  label: 'Include Group Calendars',
                  checked: this.properties.includeGroupCalendars || false,
                  onText: 'Include',
                  offText: 'Exclude',
                  disabled: !this.properties.useGraphAPI
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel('refreshIntervalLabel', {
                  text: 'Auto-refresh Settings'
                }),
                PropertyPaneSlider('refreshInterval', {
                  label: 'Auto-refresh Interval (minutes, 0 = disabled)',
                  min: 0,
                  max: AppConstants.VALIDATION.MAX_REFRESH_INTERVAL,
                  step: 5,
                  showValue: true,
                  value: this.properties.refreshInterval
                })
              ]
            },
            
            // Performance & Caching Group
            {
              groupName: 'Performance & Caching',
              groupFields: [
                PropertyPaneToggle('enableCaching', {
                  label: 'Enable Caching',
                  checked: this.properties.enableCaching !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('cacheTimeout', {
                  label: 'Cache Timeout (minutes)',
                  min: 5,
                  max: 60,
                  step: 5,
                  showValue: true,
                  value: this.properties.cacheTimeout || AppConstants.CACHE_DURATION_MINUTES,
                  disabled: !this.properties.enableCaching
                }),
                PropertyPaneButton('clearCache', {
                  text: 'Clear Cache',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this._clearCache()
                })
              ]
            }
          ]
        },
        
        // Advanced Features Page
        {
          header: {
            description: 'Configure advanced features and user experience options'
          },
          groups: [
            // Feature Toggles Group
            {
              groupName: 'Advanced Features',
              groupFields: [
                PropertyPaneToggle('enableConflictDetection', {
                  label: 'Conflict Detection',
                  checked: this.properties.enableConflictDetection !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableSearch', {
                  label: 'Search Functionality',
                  checked: this.properties.enableSearch !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableFilters', {
                  label: 'Advanced Filters',
                  checked: this.properties.enableFilters !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableExport', {
                  label: 'Export Functionality',
                  checked: this.properties.enableExport !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            },
            
            // Accessibility & UX Group
            {
              groupName: 'Accessibility & User Experience',
              groupFields: [
                PropertyPaneToggle('enableKeyboardNavigation', {
                  label: 'Keyboard Navigation',
                  checked: this.properties.enableKeyboardNavigation !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableHighContrast', {
                  label: 'High Contrast Support',
                  checked: this.properties.enableHighContrast !== false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('showLoadingAnimation', {
                  label: 'Loading Animations',
                  checked: this.properties.showLoadingAnimation !== false,
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            },
            
            // Developer Options Group
            {
              groupName: 'Developer Options',
              groupFields: [
                PropertyPaneToggle('enableDebugMode', {
                  label: 'Debug Mode',
                  checked: this.properties.enableDebugMode || false,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneDropdown('logLevel', {
                  label: 'Log Level',
                  options: [
                    { key: 'error', text: 'Error Only' },
                    { key: 'warning', text: 'Warning and Above' },
                    { key: 'info', text: 'Info and Above' },
                    { key: 'debug', text: 'All Messages' }
                  ],
                  selectedKey: this.properties.logLevel || 'error',
                  disabled: !this.properties.enableDebugMode
                }),
                PropertyPaneToggle('showPerformanceMetrics', {
                  label: 'Performance Metrics',
                  checked: this.properties.showPerformanceMetrics || false,
                  onText: 'Show',
                  offText: 'Hide',
                  disabled: !this.properties.enableDebugMode
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneButton('exportConfig', {
                  text: 'Export Configuration',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this._exportConfiguration()
                }),
                PropertyPaneLink('documentation', {
                  href: AppConstants.URLS.DOCUMENTATION,
                  text: 'View Documentation',
                  target: '_blank'
                }),
                PropertyPaneLink('support', {
                  href: AppConstants.URLS.SUPPORT,
                  text: 'Get Support',
                  target: '_blank'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected override onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    
    // Handle dependent property changes
    if (propertyPath === 'useGraphAPI') {
      if (!newValue) {
        // Disable Graph-dependent options when Graph API is disabled
        this.properties.includePersonalCalendars = false;
        this.properties.includeSharedCalendars = false;
        this.properties.includeGroupCalendars = false;
      }
    }
    
    if (propertyPath === 'enableCaching') {
      if (!newValue) {
        // Clear cache when caching is disabled
        this._clearCache();
      }
    }
    
    if (propertyPath === 'enableDebugMode') {
      if (!newValue) {
        // Reset debug-related properties when debug mode is disabled
        this.properties.showPerformanceMetrics = false;
        this.properties.logLevel = 'error';
      }
    }
    
    // Validate property values
    this._validateProperties();
    
    // Re-render if certain properties change
    const rerenderProperties = [
      'useGraphAPI', 
      'selectedCalendars', 
      'maxEvents', 
      'refreshInterval',
      'enableCaching',
      'cacheTimeout'
    ];
    
    if (rerenderProperties.includes(propertyPath)) {
      this.render();
    }
  }

  protected override get propertiesMetadata(): Record<string, { isSearchablePlainText: boolean }> {
    return {
      'title': { 'isSearchablePlainText': true },
      'description': { 'isSearchablePlainText': true }
    };
  }

  protected override onAfterResize(newWidth: number): void {
    // Handle responsive behavior if needed
    if (newWidth < 768) {
      // Mobile breakpoint - could trigger compact view
      this.properties.compactView = true;
    } else {
      // Reset compact view for larger screens if it was auto-set
      if (this.properties.compactView === undefined) {
        this.properties.compactView = false;
      }
    }
  }

  protected override get isRenderAsync(): boolean {
    return true;
  }

  protected override renderCompleted(): void {
    super.renderCompleted();
    
    // Log performance metrics if debug mode is enabled
    if (this.properties.enableDebugMode && this.properties.showPerformanceMetrics) {
      console.log('Multi-Calendar Aggregator render completed', {
        instanceId: this.context.instanceId,
        selectedCalendars: this.properties.selectedCalendars?.length || 0,
        maxEvents: this.properties.maxEvents,
        viewType: this.properties.viewType,
        renderTime: performance.now()
      });
    }
  }
}