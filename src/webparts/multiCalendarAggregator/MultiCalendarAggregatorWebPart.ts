import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MultiCalendarAggregatorWebPartStrings';
import MultiCalendarAggregator from './components/MultiCalendarAggregator';
import { IMultiCalendarAggregatorProps } from './components/IMultiCalendarAggregatorProps';
import { CalendarViewType } from './models/ICalendarModels';

export interface IMultiCalendarAggregatorWebPartProps {
  title: string;
  selectedCalendars: string[];
  viewType: CalendarViewType;
  showWeekends: boolean;
  maxEvents: number;
  refreshInterval: number;
  useGraphAPI: boolean;
  colorCoding: boolean;
}

export default class MultiCalendarAggregatorWebPart extends BaseClientSideWebPart<IMultiCalendarAggregatorWebPartProps> {

  public override render(): void {
    const element: React.ReactElement<IMultiCalendarAggregatorProps> = React.createElement(
      MultiCalendarAggregator,
      {
        title: this.properties.title,
        selectedCalendars: this.properties.selectedCalendars || [],
        viewType: this.properties.viewType as CalendarViewType,
        showWeekends: this.properties.showWeekends,
        maxEvents: this.properties.maxEvents,
        refreshInterval: this.properties.refreshInterval,
        useGraphAPI: this.properties.useGraphAPI,
        colorCoding: this.properties.colorCoding,
        isDarkTheme: false, // Default theme detection for SharePoint
        environmentMessage: this._getEnvironmentMessage(),
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected override onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected override get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected override getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: this.properties.title
                }),
                PropertyPaneDropdown('viewType', {
                  label: 'Default View',
                  options: [
                    { key: 'month', text: 'Month' },
                    { key: 'week', text: 'Week' },
                    { key: 'day', text: 'Day' },
                    { key: 'agenda', text: 'Agenda' },
                    { key: 'timeline', text: 'Timeline' }
                  ],
                  selectedKey: this.properties.viewType
                }),
                PropertyPaneSlider('maxEvents', {
                  label: 'Maximum Events',
                  min: 10,
                  max: 500,
                  step: 10,
                  showValue: true,
                  value: this.properties.maxEvents
                }),
                PropertyPaneCheckbox('showWeekends', {
                  text: 'Show Weekends',
                  checked: this.properties.showWeekends
                }),
                PropertyPaneCheckbox('colorCoding', {
                  text: 'Enable Color Coding',
                  checked: this.properties.colorCoding
                })
              ]
            },
            {
              groupName: 'Data Sources',
              groupFields: [
                PropertyPaneCheckbox('useGraphAPI', {
                  text: 'Use Microsoft Graph API (Exchange/Outlook)',
                  checked: this.properties.useGraphAPI
                }),
                PropertyPaneSlider('refreshInterval', {
                  label: 'Auto-refresh interval (minutes, 0 = disabled)',
                  min: 0,
                  max: 60,
                  step: 5,
                  showValue: true,
                  value: this.properties.refreshInterval
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
    
    // Handle property changes that might require re-rendering
    if (propertyPath === 'useGraphAPI' || propertyPath === 'selectedCalendars') {
      this.render();
    }
  }

  protected override get propertiesMetadata(): Record<string, { isSearchablePlainText: boolean }> {
    return {
      'title': {
        'isSearchablePlainText': true
      }
    };
  }

  protected override onInit(): Promise<void> {
    return this._getEnvironmentMessage().length > 0 ? Promise.resolve() : Promise.resolve();
  }
}