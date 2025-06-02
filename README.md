# Multi-Calendar Aggregator

## Summary

The Multi-Calendar Aggregator is a comprehensive SharePoint Framework (SPFx) web part that provides a unified view of calendar events from multiple sources including SharePoint calendar lists and Microsoft Exchange/Outlook calendars. This solution enables users to view, search, and manage events from different calendar sources in a single, intuitive interface.

![Multi-Calendar Aggregator](./assets/multi-calendar-hero.png)

## Features

### Core Functionality
- **📅 Multi-Source Integration**: Connect to SharePoint calendar lists and Exchange/Outlook calendars
- **🎨 Multiple View Types**: Month, week, day, agenda, and timeline views
- **🔍 Advanced Search**: Full-text search across events from all connected calendars  
- **🎯 Smart Filtering**: Filter by calendar source, category, date range, and custom criteria
- **📊 Color Coding**: Visual distinction between different calendar sources
- **📱 Responsive Design**: Optimized for desktop, tablet, and mobile devices

### Advanced Features
- **⚡ Performance Optimized**: Intelligent caching and lazy loading
- **🔄 Real-time Sync**: Automatic refresh with configurable intervals
- **📤 Export Capabilities**: Export to ICS, CSV, JSON, and Excel formats
- **🌐 Multi-tenant Support**: Works across different SharePoint sites and Exchange environments
- **♿ Accessibility**: WCAG 2.1 AA compliant with keyboard navigation and screen reader support
- **🎨 Theming**: Supports SharePoint themes and custom color schemes

### Management Features
- **⚙️ Easy Configuration**: Intuitive property panel for quick setup
- **🔐 Permission Aware**: Respects SharePoint and Exchange permissions
- **📈 Analytics**: Optional usage tracking and performance metrics
- **🔧 Admin Tools**: Bulk operations and advanced management features

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- SharePoint Online environment
- Microsoft Graph API permissions for Exchange calendar access
- Node.js v16 or higher
- SPFx development environment

## Solution

Solution|Author(s)
--------|---------
multi-calendar-aggregator | Your Name (@youralias)

## Version history

Version|Date|Comments
-------|----|--------
1.0|January 15, 2025|Initial release
1.1|January 20, 2025|Added timeline view and export features

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

1. Clone this repository
2. Ensure that you are at the solution folder
3. In the command-line run:
   - **npm install**
   - **gulp serve**

> Include any additional steps as needed.

## Configuration

### Web Part Properties

The Multi-Calendar Aggregator web part can be configured through the property panel with the following options:

#### Basic Settings
| Property | Type | Description | Default |
|----------|------|-------------|---------|
| Title | string | Display title for the web part | "Multi-Calendar Aggregator" |
| Default View | Choice | Initial calendar view (month/week/day/agenda/timeline) | Month |
| Max Events | number | Maximum number of events to display | 100 |
| Show Weekends | boolean | Include weekends in calendar views | true |
| Color Coding | boolean | Enable color coding by calendar source | true |

#### Data Sources
| Property | Type | Description | Default |
|----------|------|-------------|---------|
| Use Graph API | boolean | Enable Microsoft Graph for Exchange calendars | true |
| Selected Calendars | string[] | Array of calendar IDs to include | [] |
| Refresh Interval | number | Auto-refresh interval in minutes (0 = disabled) | 15 |

#### Display Options
| Property | Type | Description | Default |
|----------|------|-------------|---------|
| Enable Search | boolean | Show search functionality | true |
| Enable Filters | boolean | Show filtering options | true |
| Enable Export | boolean | Show export options | true |

### Microsoft Graph Permissions

To enable Exchange/Outlook calendar integration, the following Microsoft Graph permissions are required:

#### Application Permissions (for tenant-wide deployment)
```json
{
  "resource": "Microsoft Graph",
  "scope": "Calendars.Read"
}
```

#### Delegated Permissions (for user-specific access)
```json
{
  "resource": "Microsoft Graph",
  "scope": "Calendars.Read Calendars.Read.Shared"
}
```

### Deployment Guide

#### 1. Package the Solution
```bash
# Build and bundle the solution
gulp build
gulp bundle --ship

# Package the solution
gulp package-solution --ship
```

#### 2. Deploy to SharePoint App Catalog
1. Upload the `.sppkg` file to your tenant or site collection app catalog
2. Deploy the solution and trust it when prompted
3. If using Microsoft Graph, approve the API permissions in SharePoint Admin Center

#### 3. Grant API Permissions (if using Graph)
1. Go to SharePoint Admin Center > Advanced > API access
2. Approve pending requests for Microsoft Graph permissions
3. Ensure users have appropriate Exchange/Outlook access

#### 4. Add to SharePoint Pages
1. Edit a SharePoint page
2. Add the "Multi-Calendar Aggregator" web part
3. Configure the web part properties as needed

## Features Deep Dive

### 📅 Calendar Views

#### Month View
- Traditional calendar grid layout
- Color-coded events by source
- All-day events displayed prominently
- Click events for detailed view

#### Week View  
- 7-day horizontal layout
- Time slots with hourly precision
- Overlapping event detection
- Current time indicator

#### Day View
- Single day detailed view
- 15-minute time increments
- Ideal for busy schedules
- Real-time current time marker

#### Agenda View
- List-based event display
- Grouping by date, calendar, or category
- Configurable date range (7/14/30 days)
- Compact view for quick scanning

#### Timeline View
- Horizontal timeline representation
- Drag and scroll navigation
- Multiple date ranges (day/week/month)
- Visual event duration display

### 🔍 Search and Filtering

#### Search Capabilities
- Full-text search across event titles, descriptions, and locations
- Real-time search results
- Search term highlighting
- Recent search suggestions

#### Advanced Filtering
- **Date Range**: Custom start/end dates or preset ranges
- **Calendar Sources**: Select specific calendars to include/exclude
- **Categories**: Filter by event categories
- **Event Types**: All-day, recurring, meetings, appointments
- **Attendees**: Filter by organizer or attendee names
- **Importance**: High, normal, low priority events

### 📊 Data Integration

#### SharePoint Calendars
- Automatic discovery of calendar lists
- Support for custom calendar columns
- Recurrence pattern recognition
- Permission-aware access

#### Exchange/Outlook Calendars  
- Personal calendar access via Microsoft Graph
- Shared calendar support
- Room and resource calendars
- Free/busy information

#### Caching Strategy
- Intelligent multi-level caching
- Configurable cache duration
- Background refresh capabilities
- Cache invalidation on data changes

### 🎨 Customization

#### Theming Support
- Automatic SharePoint theme detection
- Custom color palette support
- High contrast mode compatibility
- Dark mode support

#### Responsive Design
- Mobile-first approach
- Touch-friendly interface
- Adaptive layouts for different screen sizes
- Progressive enhancement

## API Reference

### CalendarService

The main service class for calendar operations:

```typescript
// Initialize the service
const calendarService = new CalendarService(context);
await calendarService.initialize();

// Get available calendar sources
const sources = await calendarService.getCalendarSources(includeExchange);

// Get events from specific sources
const events = await calendarService.getEventsFromSources(sources, maxEvents);

// Search events
const results = await calendarService.searchEvents(sources, query, maxResults);
```

### Event Data Model

```typescript
interface ICalendarEvent {
  id: string;
  title: string;
  description: string;
  start: Date;
  end: Date;
  location?: string;
  category?: string;
  isAllDay: boolean;
  isRecurring: boolean;
  calendarId: string;
  calendarTitle: string;
  calendarType: CalendarSourceType;
  organizer: string;
  attendees?: IEventAttendee[];
  // ... additional properties
}
```

## Development

### Setting up Development Environment

1. **Install Prerequisites**
   ```bash
   # Install SPFx globally
   npm install -g @microsoft/generator-sharepoint
   
   # Install Yeoman and Gulp
   npm install -g yo gulp
   ```

2. **Clone and Setup**
   ```bash
   git clone [repository-url]
   cd multi-calendar-aggregator
   npm install
   ```

3. **Development Server**
   ```bash
   # Start development server
   gulp serve
   
   # With specific tenant
   gulp serve --config=your-tenant-config
   ```

### Project Structure

```
src/
├── webparts/
│   └── multiCalendarAggregator/
│       ├── components/           # React components
│       │   ├── MultiCalendarAggregator.tsx
│       │   ├── CalendarSourcesPanel.tsx
│       │   ├── EventDetailsPanel.tsx
│       │   ├── AgendaView.tsx
│       │   └── TimelineView.tsx
│       ├── services/            # Data services
│       │   ├── CalendarService.ts
│       │   ├── SharePointCalendarService.ts
│       │   ├── ExchangeCalendarService.ts
│       │   └── CacheService.ts
│       ├── models/              # Type definitions
│       │   ├── ICalendarModels.ts
│       │   ├── IEventModels.ts
│       │   └── IConfigurationModels.ts
│       ├── utils/               # Utility functions
│       │   ├── DateUtils.ts
│       │   ├── ValidationUtils.ts
│       │   ├── ColorUtils.ts
│       │   └── ExportUtils.ts
│       └── constants/           # Application constants
│           └── AppConstants.ts
```

### Building for Production

```bash
# Clean previous builds
gulp clean

# Build for production
gulp build --ship

# Bundle assets
gulp bundle --ship

# Create solution package
gulp package-solution --ship
```

### Testing

```bash
# Run unit tests
npm test

# Run with coverage
npm run test:coverage

# Run end-to-end tests
npm run test:e2e
```

## Troubleshooting

### Common Issues

#### 1. Microsoft Graph Permissions
**Problem**: Exchange calendars not loading
**Solution**: 
- Verify Graph permissions are approved in SharePoint Admin Center
- Check user has appropriate Exchange/Outlook licenses
- Ensure tenant admin has granted consent

#### 2. SharePoint Calendar Access
**Problem**: SharePoint calendars not appearing
**Solution**:
- Verify user has read access to calendar lists
- Check if calendar lists exist in accessible sites
- Ensure web part has necessary SharePoint permissions

#### 3. Performance Issues
**Problem**: Slow loading with many calendars
**Solution**:
- Reduce max events limit in web part properties
- Enable caching in advanced settings
- Limit number of selected calendar sources

#### 4. Display Issues
**Problem**: Events not showing correctly
**Solution**:
- Check browser compatibility (requires modern browsers)
- Verify SharePoint theme compatibility
- Clear browser cache and reload

### Debug Mode

Enable debug mode by adding this to your page URL:
```
?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

### Support

For additional support:
1. Check the [Issues](../../issues) section for known problems
2. Review SharePoint Framework documentation
3. Contact your SharePoint administrator for permission-related issues

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Workflow
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

### Code Standards
- Follow TypeScript best practices
- Use ESLint and Prettier for code formatting
- Write unit tests for new features
- Document public APIs
- Follow semantic versioning

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<img src="https://m365-visitor-stats.azurewebsites.net/sp-dev-fx-webparts/samples/react-multi-calendar-aggregator" />