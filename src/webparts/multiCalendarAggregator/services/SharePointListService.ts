import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { 
  ISharePointListSource, 
  ISharePointEvent, 
  CalendarSourceType, 
  ISharePointCustomField,
  ISharePointPermissions,
  IFieldMapping,
  ISharePointAttachment,
  ISharePointListConfiguration
} from '../models/ICalendarModels';
import { DateUtils } from '../utils/DateUtils';
import { ColorUtils } from '../utils/ColorUtils';

export class SharePointListService {
  private context: WebPartContext;
  private configuration: ISharePointListConfiguration;

  constructor(context: WebPartContext, configuration?: ISharePointListConfiguration) {
    this.context = context;
    this.configuration = configuration || this.getDefaultConfiguration();
  }

  /**
   * Get default configuration
   */
  private getDefaultConfiguration(): ISharePointListConfiguration {
    return {
      includeCustomLists: true,
      includeTaskLists: true,
      includeIssueLists: true,
      includeAnnouncementLists: true,
      autoDiscoverDateFields: true,
      enableWorkflowIntegration: false,
      enableVersionHistory: false,
      enableComments: false,
      maxItemsPerList: 1000,
      dateRangeMonths: 6
    };
  }

  /**
   * Discover SharePoint lists that can be used as calendar sources
   */
  public async discoverCalendarLists(siteUrl?: string): Promise<ISharePointListSource[]> {
    const targetSiteUrl = siteUrl || this.context.pageContext.web.absoluteUrl;
    const lists: ISharePointListSource[] = [];

    try {
      console.log(`Discovering calendar lists in: ${targetSiteUrl}`);
      
      // Get all potential calendar lists
      const filter = this.buildListFilter();
      const select = this.buildListSelect();
      const expand = this.buildListExpand();

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${targetSiteUrl}/_api/web/lists?${filter}&${select}&${expand}`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const listData = data.d?.results || data.value || [];

        console.log(`Found ${listData.length} potential lists`);

        for (const list of listData) {
          try {
            const listSource = await this.mapToListSource(list, targetSiteUrl);
            if (listSource) {
              lists.push(listSource);
              console.log(`Mapped list: ${listSource.title} (${listSource.listType})`);
            }
          } catch (error) {
            console.warn(`Error processing list ${list.Title}:`, error);
          }
        }
      } else {
        console.error(`Failed to get lists. Status: ${response.status}`);
      }
    } catch (error) {
      console.error('Error discovering SharePoint lists:', error);
      throw new Error(`Failed to discover SharePoint lists: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }

    console.log(`Successfully discovered ${lists.length} calendar lists`);
    return lists;
  }

  /**
   * Build list filter query
   */
  private buildListFilter(): string {
    const filters: string[] = [
      'Hidden eq false',
      'IsCatalog eq false'
    ];

    // Base templates to include
    const templates: number[] = [100]; // Generic list (always include)

    if (this.configuration.includeCustomLists) {
      templates.push(106); // Events list
    }

    if (this.configuration.includeTaskLists) {
      templates.push(107); // Tasks list
    }

    if (this.configuration.includeIssueLists) {
      templates.push(171); // Issue tracking list
    }

    if (this.configuration.includeAnnouncementLists) {
      templates.push(104); // Announcements list
    }

    const templateFilter = `(${templates.map(t => `BaseTemplate eq ${t}`).join(' or ')})`;
    filters.push(templateFilter);

    return `$filter=${filters.join(' and ')}`;
  }

  /**
   * Build list select query
   */
  private buildListSelect(): string {
    const fields = [
      'Id',
      'Title', 
      'Description',
      'BaseTemplate',
      'ItemCount',
      'LastItemModifiedDate',
      'Created',
      'DefaultViewUrl',
      'EnableVersioning',
      'EnableModeration',
      'HasUniqueRoleAssignments',
      'WorkflowAssociations',
      'ContentTypesEnabled',
      'RootFolder/ServerRelativeUrl'
    ];

    return `$select=${fields.join(',')}`;
  }

  /**
   * Build list expand query
   */
  private buildListExpand(): string {
    const expands = [
      'Fields($filter=Hidden eq false and ReadOnlyField eq false)',
      'Views($filter=Hidden eq false)',
      'RootFolder',
      'ContentTypes($filter=Hidden eq false)',
      'WorkflowAssociations'
    ];

    return `$expand=${expands.join(',')}`;
  }

  /**
   * Map SharePoint list to calendar source
   */
  private async mapToListSource(listData: any, siteUrl: string): Promise<ISharePointListSource | null> {
    try {
      // Determine list type based on template and fields
      const listType = this.determineListType(listData);
      if (!listType) {
        console.log(`Skipping list ${listData.Title}: not a supported calendar list type`);
        return null;
      }

      // Check if list has date fields that could be used for events
      const fields = listData.Fields?.results || [];
      const dateFields = this.findDateFields(fields);
      
      if (dateFields.length === 0 && this.configuration.autoDiscoverDateFields) {
        console.log(`Skipping list ${listData.Title}: no date fields found`);
        return null;
      }

      // Get detailed permissions for the list
      const permissions = await this.getListPermissions(listData.Id, siteUrl);
      if (!permissions.canRead) {
        console.log(`Skipping list ${listData.Title}: insufficient permissions`);
        return null;
      }

      // Map custom fields
      const customFields = this.mapCustomFields(fields);
      
      // Create field mappings
      const fieldMappings = this.createFieldMappings(listType, fields);

      const listSource: ISharePointListSource = {
        id: listData.Id,
        title: listData.Title,
        description: listData.Description || `SharePoint ${listType} list with ${listData.ItemCount || 0} items`,
        type: CalendarSourceType.SharePointList,
        listType,
        listTemplate: listData.BaseTemplate,
        url: `${siteUrl}${listData.DefaultViewUrl}`,
        viewUrl: `${siteUrl}${listData.DefaultViewUrl}`,
        siteTitle: this.extractSiteTitle(siteUrl),
        siteUrl: siteUrl,
        itemCount: listData.ItemCount || 0,
        lastModified: listData.LastItemModifiedDate,
        color: this.getListTypeColor(listType),
        isEnabled: true,
        canEdit: permissions.canWrite,
        canShare: permissions.canManagePermissions,
        customFields,
        permissions,
        workflowEnabled: (listData.WorkflowAssociations?.results || []).length > 0,
        versioningEnabled: listData.EnableVersioning,
        contentTypesEnabled: listData.ContentTypesEnabled,
        fieldMappings
      };

      return listSource;
    } catch (error) {
      console.error('Error mapping list to source:', error);
      return null;
    }
  }

  /**
   * Extract site title from URL
   */
  private extractSiteTitle(siteUrl: string): string {
    try {
      const urlParts = siteUrl.split('/');
      return urlParts[urlParts.length - 1] || 'SharePoint Site';
    } catch {
      return 'SharePoint Site';
    }
  }

  /**
   * Determine list type based on template and fields
   */
  private determineListType(listData: any): 'Events' | 'Calendar' | 'Custom' | 'Tasks' | 'Issues' | 'Announcements' | null {
    const template = listData.BaseTemplate;
    const fields = listData.Fields?.results || [];
    const title = listData.Title?.toLowerCase() || '';
    
    switch (template) {
      case 106: return 'Events';
      case 107: return 'Tasks';
      case 171: return 'Issues';
      case 104: return 'Announcements';
      case 100: {
        // Generic list - determine by fields and title
        const hasEventDate = fields.some((f: any) => f.InternalName === 'EventDate');
        const hasStartDate = fields.some((f: any) => f.InternalName === 'StartDate');
        const hasDueDate = fields.some((f: any) => f.InternalName === 'DueDate');
        const hasDateFields = fields.some((f: any) => f.TypeAsString === 'DateTime');
        
        // Check title for hints
        if (title.includes('event') || title.includes('calendar')) {
          return hasEventDate ? 'Events' : 'Custom';
        }
        if (title.includes('task') || title.includes('todo')) {
          return 'Tasks';
        }
        if (title.includes('issue') || title.includes('ticket')) {
          return 'Issues';
        }
        if (title.includes('announcement') || title.includes('news')) {
          return 'Announcements';
        }
        
        // Determine by fields
        if (hasEventDate) return 'Events';
        if (hasStartDate || hasDueDate) return 'Tasks';
        if (hasDateFields) return 'Custom';
        
        return null;
      }
      default: return null;
    }
  }

  /**
   * Find date fields in the list
   */
  private findDateFields(fields: any[]): any[] {
    return fields.filter(field => {
      const isDateTimeField = field.TypeAsString === 'DateTime';
      const isKnownDateField = [
        'EventDate', 'EndDate', 'StartDate', 'DueDate', 
        'Created', 'Modified', 'PublishedDate', 'ExpirationDate'
      ].includes(field.InternalName);
      
      return isDateTimeField || isKnownDateField;
    });
  }

  /**
   * Create field mappings based on list type and available fields
   */
  private createFieldMappings(listType: string, fields: any[]): IFieldMapping {
    const fieldMap: IFieldMapping = {
      startDateField: 'Created' // Default fallback
    };

    const fieldNames = fields.map((f: any) => f.InternalName);

    switch (listType) {
      case 'Events': {
        fieldMap.startDateField = 'EventDate';
        fieldMap.endDateField = 'EndDate';
        fieldMap.titleField = 'Title';
        fieldMap.descriptionField = 'Description';
        fieldMap.locationField = 'Location';
        fieldMap.categoryField = 'Category';
        fieldMap.allDayField = 'fAllDayEvent';
        fieldMap.recurrenceField = 'fRecurrence';
        break;
      }
      case 'Tasks': {
        fieldMap.startDateField = 'StartDate';
        fieldMap.endDateField = 'DueDate';
        fieldMap.titleField = 'Title';
        fieldMap.descriptionField = 'Description';
        fieldMap.statusField = 'Status';
        fieldMap.organizerField = 'AssignedTo';
        fieldMap.importanceField = 'Priority';
        break;
      }
      case 'Issues': {
        fieldMap.startDateField = 'Created';
        fieldMap.endDateField = 'DueDate';
        fieldMap.titleField = 'Title';
        fieldMap.descriptionField = 'Description';
        fieldMap.statusField = 'Status';
        fieldMap.organizerField = 'AssignedTo';
        fieldMap.importanceField = 'Priority';
        fieldMap.categoryField = 'Category';
        break;
      }
      case 'Announcements': {
        fieldMap.startDateField = 'Created';
        fieldMap.endDateField = 'Expires';
        fieldMap.titleField = 'Title';
        fieldMap.descriptionField = 'Body';
        break;
      }
      default: {
        // Custom list - try to find appropriate fields
        if (fieldNames.includes('EventDate')) fieldMap.startDateField = 'EventDate';
        else if (fieldNames.includes('StartDate')) fieldMap.startDateField = 'StartDate';
        else if (fieldNames.includes('DueDate')) fieldMap.startDateField = 'DueDate';
        
        if (fieldNames.includes('EndDate')) fieldMap.endDateField = 'EndDate';
        else if (fieldNames.includes('DueDate')) fieldMap.endDateField = 'DueDate';
        
        fieldMap.titleField = 'Title';
        if (fieldNames.includes('Description')) fieldMap.descriptionField = 'Description';
        if (fieldNames.includes('Location')) fieldMap.locationField = 'Location';
        if (fieldNames.includes('Category')) fieldMap.categoryField = 'Category';
        break;
      }
    }

    return fieldMap;
  }

  /**
   * Get list permissions for current user
   */
  private async getListPermissions(listId: string, siteUrl: string): Promise<ISharePointPermissions> {
    try {
      const response = await this.context.spHttpClient.get(
        `${siteUrl}/_api/web/lists(guid'${listId}')/EffectiveBasePermissions`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const permissions = data.d?.EffectiveBasePermissions || data.EffectiveBasePermissions;
        
        return {
          canRead: true, // If we can query, we can read
          canWrite: this.hasPermission(permissions, 'EditListItems'),
          canDelete: this.hasPermission(permissions, 'DeleteListItems'),
          canManagePermissions: this.hasPermission(permissions, 'ManagePermissions'),
          canManageViews: this.hasPermission(permissions, 'ManageViews'),
          canApprove: this.hasPermission(permissions, 'ApproveItems'),
          canManageWebParts: this.hasPermission(permissions, 'AddAndCustomizePages'),
          effectivePermissions: permissions ? [permissions.High, permissions.Low] : [],
          permissionLevel: this.getPermissionLevel(permissions)
        };
      }
    } catch (error) {
      console.warn('Could not get list permissions:', error);
    }

    return {
      canRead: true,
      canWrite: false,
      canDelete: false,
      canManagePermissions: false,
      canManageViews: false,
      canApprove: false,
      canManageWebParts: false,
      permissionLevel: 'Read'
    };
  }

  /**
   * Check if user has specific permission
   */
  private hasPermission(permissions: any, permissionName: string): boolean {
    if (!permissions) return false;
    
    // SharePoint permission bit checking (simplified)
    // In production, implement proper SP permission bit calculations
    const permissionMasks: { [key: string]: { high: number; low: number } } = {
      'EditListItems': { high: 0, low: 2 },
      'DeleteListItems': { high: 0, low: 8 },
      'ManagePermissions': { high: 0, low: 1073741824 },
      'ManageViews': { high: 0, low: 256 },
      'ApproveItems': { high: 0, low: 16 },
      'AddAndCustomizePages': { high: 0, low: 4194304 }
    };

    const mask = permissionMasks[permissionName];
    if (!mask) return false;

    const high = permissions.High || 0;
    const low = permissions.Low || 0;

    return (high & mask.high) === mask.high && (low & mask.low) === mask.low;
  }

  /**
   * Get permission level description
   */
  private getPermissionLevel(permissions: any): string {
    if (!permissions) return 'None';
    
    if (this.hasPermission(permissions, 'ManagePermissions')) return 'Full Control';
    if (this.hasPermission(permissions, 'EditListItems')) return 'Contribute';
    return 'Read';
  }

  /**
   * Map custom fields from SharePoint list
   */
  private mapCustomFields(fields: any[]): ISharePointCustomField[] {
    const systemFields = [
      'ID', 'Title', 'Created', 'Modified', 'Author', 'Editor', 'ContentType',
      'Attachments', 'WorkflowVersion', 'WorkflowInstanceID', '_UIVersionString',
      'FileSystemObjectType', 'ServerRedirectedEmbedUri', 'ServerRedirectedEmbedUrl',
      'ContentTypeId', 'ComplianceAssetId', '_ComplianceFlags', '_ComplianceTag',
      '_ComplianceTagWrittenTime', '_ComplianceTagUserId', '_IsRecord'
    ];

    return fields
.filter((field: any) => 
        !field.Hidden && 
        !field.ReadOnlyField && 
        !systemFields.includes(field.InternalName) &&
        !field.InternalName.startsWith('_')
      )
      .map(field => ({
        internalName: field.InternalName,
        displayName: field.Title,
        fieldType: this.mapSharePointFieldType(field.TypeAsString),
        required: field.Required,
        mappedTo: this.getFieldMapping(field.InternalName),
        choices: field.Choices?.results || undefined,
        defaultValue: field.DefaultValue || undefined,
        isMultiValue: field.TypeAsString?.includes('Multi') || false
      }));
  }

  /**
   * Map SharePoint field types to our simplified types
   */
  private mapSharePointFieldType(spFieldType: string): ISharePointCustomField['fieldType'] {
    switch (spFieldType) {
      case 'DateTime': return 'DateTime';
      case 'Choice': return 'Choice';
      case 'MultiChoice': return 'MultiChoice';
      case 'User':
      case 'UserMulti': return 'User';
      case 'Lookup':
      case 'LookupMulti': return 'Lookup';
      case 'Number':
      case 'Currency': return 'Number';
      case 'Boolean': return 'Boolean';
      case 'URL': return 'URL';
      case 'Note': return 'Note';
      default: return 'Text';
    }
  }

  /**
   * Get suggested field mapping based on field name
   */
  private getFieldMapping(internalName: string): ISharePointCustomField['mappedTo'] {
    const name = internalName.toLowerCase();
    
    if (name.includes('title') || name.includes('subject')) return 'title';
    if (name.includes('description') || name.includes('comment') || name.includes('body') || name.includes('note')) return 'description';
    if (name.includes('location') || name.includes('room') || name.includes('venue') || name.includes('where')) return 'location';
    if (name.includes('category') || name.includes('type') || name.includes('priority')) return 'category';
    if (name.includes('organizer') || name.includes('owner') || name.includes('responsible')) return 'organizer';
    if (name.includes('attendee') || name.includes('participant') || name.includes('assigned')) return 'attendees';
    
    return 'custom';
  }

  /**
   * Get color based on list type
   */
  private getListTypeColor(listType: string): string {
    const colors = {
      'Events': '#0078d4',
      'Calendar': '#0078d4', 
      'Tasks': '#107c10',
      'Issues': '#d13438',
      'Announcements': '#ff8c00',
      'Custom': '#881798'
    };
    return colors[listType as keyof typeof colors] || ColorUtils.generateColorFromString(listType);
  }

  /**
   * Get events from SharePoint list
   */
  public async getEventsFromList(source: ISharePointListSource, maxEvents: number = 100): Promise<ISharePointEvent[]> {
    const events: ISharePointEvent[] = [];

    try {
      console.log(`Getting events from list: ${source.title}`);
      
      // Build field selection based on list type and custom fields
      const selectFields = this.buildSelectFields(source);
      const filterQuery = this.buildFilterQuery(source);
      const orderBy = this.getOrderByField(source);
      
      const apiUrl = `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items?` +
        `$select=${selectFields}&` +
        `$expand=AttachmentFiles,Author,Editor,AssignedTo&` +
        `$filter=${filterQuery}&` +
        `$orderby=${orderBy}&` +
        `$top=${Math.min(maxEvents, this.configuration.maxItemsPerList)}`;

      console.log(`API URL: ${apiUrl}`);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.d?.results || data.value || [];

        console.log(`Retrieved ${items.length} items from list`);
        for (const item of items) {
         try {
           const event = this.mapListItemToEvent(item, source);
           if (event) {
             events.push(event);
           }
         } catch (error) {
           console.warn(`Error mapping item ${item.ID} to event:`, error);
         }
       }

       // Get additional data if enabled
       if (this.configuration.enableVersionHistory || this.configuration.enableComments) {
         await this.enrichEventsWithAdditionalData(events, source);
       }

     } else {
       throw new Error(`Failed to fetch events. Status: ${response.status}`);
     }
   } catch (error) {
     console.error(`Error getting events from list ${source.title}:`, error);
     throw error;
   }

   console.log(`Successfully retrieved ${events.length} events from list: ${source.title}`);
   return events;
 }

 /**
  * Build select fields query based on list configuration
  */
 private buildSelectFields(source: ISharePointListSource): string {
   const baseFields = [
     'ID', 'Title', 'Created', 'Modified', 'Author/Title', 'Author/EMail', 
     'Editor/Title', 'Editor/EMail', 'HasAttachments', 'ContentType/Name'
   ];
   
   // Add fields based on field mappings
   if (source.fieldMappings) {
     const mappings = source.fieldMappings;
     if (mappings.startDateField) baseFields.push(mappings.startDateField);
     if (mappings.endDateField) baseFields.push(mappings.endDateField);
     if (mappings.descriptionField) baseFields.push(mappings.descriptionField);
     if (mappings.locationField) baseFields.push(mappings.locationField);
     if (mappings.categoryField) baseFields.push(mappings.categoryField);
     if (mappings.organizerField) baseFields.push(`${mappings.organizerField}/Title`, `${mappings.organizerField}/EMail`);
     if (mappings.allDayField) baseFields.push(mappings.allDayField);
     if (mappings.recurrenceField) baseFields.push(mappings.recurrenceField);
     if (mappings.importanceField) baseFields.push(mappings.importanceField);
     if (mappings.statusField) baseFields.push(mappings.statusField);
   }

   // Add list type specific fields
   switch (source.listType) {
     case 'Events':
       baseFields.push('EventDate', 'EndDate', 'Location', 'Description', 'Category', 'fAllDayEvent', 'fRecurrence');
       break;
     case 'Tasks':
       baseFields.push('StartDate', 'DueDate', 'Status', 'Priority', 'AssignedTo/Title', 'AssignedTo/EMail', 'PercentComplete');
       break;
     case 'Issues':
       baseFields.push('DueDate', 'Status', 'Priority', 'AssignedTo/Title', 'AssignedTo/EMail', 'Category', 'IssueStatus');
       break;
     case 'Announcements':
       baseFields.push('Body', 'Expires', 'Priority');
       break;
   }

   // Add custom fields
   source.customFields?.forEach(field => {
     if (field.fieldType === 'User') {
       baseFields.push(`${field.internalName}/Title`, `${field.internalName}/EMail`);
     } else {
       baseFields.push(field.internalName);
     }
   });

   // Remove duplicates
   return [...new Set(baseFields)].join(',');
 }

 /**
  * Build filter query for date-based filtering
  */
 private buildFilterQuery(source: ISharePointListSource): string {
   const now = new Date();
   const futureDate = new Date();
   futureDate.setMonth(futureDate.getMonth() + this.configuration.dateRangeMonths);

   const startDateField = source.fieldMappings?.startDateField || this.getDefaultStartDateField(source);
   
   return `${startDateField} ge datetime'${now.toISOString()}' and ${startDateField} le datetime'${futureDate.toISOString()}'`;
 }

 /**
  * Get the default start date field for filtering and ordering
  */
 private getDefaultStartDateField(source: ISharePointListSource): string {
   if (source.fieldMappings?.startDateField) {
     return source.fieldMappings.startDateField;
   }

   switch (source.listType) {
     case 'Events': return 'EventDate';
     case 'Tasks': return 'StartDate';
     case 'Issues': return 'Created';
     case 'Announcements': return 'Created';
     default: {
       // For custom lists, find the first date field
       const dateField = source.customFields?.find(f => f.fieldType === 'DateTime');
       return dateField?.internalName || 'Created';
     }
   }
 }

 /**
  * Get order by field
  */
 private getOrderByField(source: ISharePointListSource): string {
   const startDateField = source.fieldMappings?.startDateField || this.getDefaultStartDateField(source);
   return `${startDateField} asc`;
 }

 /**
  * Map SharePoint list item to calendar event
  */
 private mapListItemToEvent(item: any, source: ISharePointListSource): ISharePointEvent | null {
   try {
     const startDate = this.getItemStartDate(item, source);
     if (!startDate) {
       console.warn(`Item ${item.ID} has no valid start date`);
       return null;
     }

     const endDate = this.getItemEndDate(item, source, startDate);
     
     const event: ISharePointEvent = {
       id: `sp_list_${source.id}_${item.ID}`,
       title: this.getItemTitle(item, source),
       description: this.getItemDescription(item, source),
       start: startDate,
       end: endDate,
       location: this.getItemLocation(item, source),
       category: this.getItemCategory(item, source),
       isAllDay: this.getIsAllDay(item, source),
       isRecurring: this.getIsRecurring(item, source),
       calendarId: source.id,
       calendarTitle: source.title,
       calendarType: source.type,
       organizer: this.getItemOrganizer(item, source),
       created: new Date(item.Created),
       modified: new Date(item.Modified),
       webUrl: this.getItemWebUrl(item, source),
       color: source.color,
       importance: this.getItemImportance(item, source),
       
       // SharePoint-specific properties
       listItemId: item.ID,
       listItemUrl: this.getItemWebUrl(item, source),
       workflowStatus: item.WorkflowInstanceID ? 'Active' : 'None',
       approvalStatus: this.getApprovalStatus(item),
       customFields: this.extractCustomFields(item, source),
       sharePointAttachments: this.mapAttachments(item.AttachmentFiles?.results || []),
       contentType: item.ContentType?.Name || 'Item',
       etag: item.__metadata?.etag,
       hasAttachments: item.HasAttachments || false,
       percentComplete: item.PercentComplete || 0,
       assignedTo: this.getAssignedTo(item),
       priority: item.Priority || 'Normal',
       taskStatus: item.Status,
       publishedDate: item.PublishedDate ? new Date(item.PublishedDate) : undefined,
       expirationDate: item.Expires ? new Date(item.Expires) : undefined
     };

     return event;
   } catch (error) {
     console.warn(`Error mapping list item ${item.ID} to event:`, error);
     return null;
   }
 }

 /**
  * Get start date from list item
  */
 private getItemStartDate(item: any, source: ISharePointListSource): Date | null {
   const dateField = source.fieldMappings?.startDateField || this.getDefaultStartDateField(source);
   const dateValue = item[dateField];
   
   if (!dateValue) return null;
   
   try {
     return new Date(dateValue);
   } catch {
     return null;
   }
 }

 /**
  * Get end date from list item
  */
 private getItemEndDate(item: any, source: ISharePointListSource, startDate: Date): Date {
   const endDateField = source.fieldMappings?.endDateField;
   
   if (endDateField && item[endDateField]) {
     try {
       const endDate = new Date(item[endDateField]);
       if (endDate > startDate) {
         return endDate;
       }
     } catch {
       // Fall through to default calculation
     }
   }

   // Default end date calculation based on list type
   switch (source.listType) {
     case 'Events': {
       const endDate = item.EndDate ? new Date(item.EndDate) : null;
       if (endDate && endDate > startDate) {
         return endDate;
       }
       return new Date(startDate.getTime() + 60 * 60 * 1000); // Default 1 hour
     }
     case 'Tasks': {
       const dueDate = item.DueDate ? new Date(item.DueDate) : null;
       if (dueDate && dueDate > startDate) {
         return dueDate;
       }
       return new Date(startDate.getTime() + 24 * 60 * 60 * 1000); // Default 1 day
     }
     case 'Announcements': {
       const expires = item.Expires ? new Date(item.Expires) : null;
       if (expires && expires > startDate) {
         return expires;
       }
       return new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000); // Default 1 week
     }
     default: {
       return new Date(startDate.getTime() + 60 * 60 * 1000); // Default 1 hour
     }
   }
 }

 /**
  * Get item title
  */
 private getItemTitle(item: any, source: ISharePointListSource): string {
   const titleField = source.fieldMappings?.titleField || 'Title';
   return item[titleField] || item.Title || 'Untitled Item';
 }

 /**
  * Get item description
  */
 private getItemDescription(item: any, source: ISharePointListSource): string {
   const descField = source.fieldMappings?.descriptionField;
   
   if (descField && item[descField]) {
     return this.stripHtml(item[descField]);
   }

   // Fallback to common description fields
   const commonFields = ['Description', 'Body', 'Comments', 'Notes'];
   for (const field of commonFields) {
     if (item[field]) {
       return this.stripHtml(item[field]);
     }
   }

   return '';
 }

 /**
  * Strip HTML tags from text
  */
 private stripHtml(html: string): string {
   if (!html || typeof html !== 'string') return '';
   
   return html
     .replace(/<[^>]*>/g, '') // Remove HTML tags
     .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
     .replace(/&amp;/g, '&') // Replace encoded ampersands
     .replace(/&lt;/g, '<') // Replace encoded less than
     .replace(/&gt;/g, '>') // Replace encoded greater than
     .replace(/&quot;/g, '"') // Replace encoded quotes
     .trim();
 }

 /**
  * Get item location
  */
 private getItemLocation(item: any, source: ISharePointListSource): string {
   const locationField = source.fieldMappings?.locationField;
   
   if (locationField && item[locationField]) {
     return item[locationField];
   }

   return item.Location || '';
 }

 /**
  * Get item category
  */
 private getItemCategory(item: any, source: ISharePointListSource): string {
   const categoryField = source.fieldMappings?.categoryField;
   
   if (categoryField && item[categoryField]) {
     return item[categoryField];
   }

   return item.Category || item.Priority || source.listType;
 }

 /**
  * Get item organizer
  */
 private getItemOrganizer(item: any, source: ISharePointListSource): string {
   const organizerField = source.fieldMappings?.organizerField;
   
   if (organizerField && item[organizerField]) {
     const organizer = item[organizerField];
     if (typeof organizer === 'object' && organizer.Title) {
       return organizer.Title;
     }
     if (typeof organizer === 'string') {
       return organizer;
     }
   }

   // Fallback to common organizer fields
   if (item.AssignedTo?.Title) return item.AssignedTo.Title;
   if (item.Author?.Title) return item.Author.Title;
   
   return 'Unknown';
 }

 /**
  * Get assigned to users
  */
 private getAssignedTo(item: any): string[] {
   const assignedTo: string[] = [];
   
   if (item.AssignedTo) {
     if (Array.isArray(item.AssignedTo)) {
       assignedTo.push(...item.AssignedTo.map((user: any) => user.Title || user));
     } else if (item.AssignedTo.Title) {
       assignedTo.push(item.AssignedTo.Title);
     }
   }
   
   return assignedTo;
 }

 /**
  * Get item importance
  */
 private getItemImportance(item: any, source: ISharePointListSource): 'low' | 'normal' | 'high' {
   const importanceField = source.fieldMappings?.importanceField;
   
   let importance = '';
   if (importanceField && item[importanceField]) {
     importance = item[importanceField].toLowerCase();
   } else if (item.Priority) {
     importance = item.Priority.toLowerCase();
   }

   switch (importance) {
     case 'high':
     case 'urgent':
     case 'critical':
     case '1':
       return 'high';
     case 'low':
     case '3':
       return 'low';
     default:
       return 'normal';
   }
 }

 /**
  * Get if item is all day
  */
 private getIsAllDay(item: any, source: ISharePointListSource): boolean {
   const allDayField = source.fieldMappings?.allDayField;
   
   if (allDayField && item[allDayField] !== undefined) {
     return Boolean(item[allDayField]);
   }

   return item.fAllDayEvent || false;
 }

 /**
  * Get if item is recurring
  */
 private getIsRecurring(item: any, source: ISharePointListSource): boolean {
   const recurrenceField = source.fieldMappings?.recurrenceField;
   
   if (recurrenceField && item[recurrenceField] !== undefined) {
     return Boolean(item[recurrenceField]);
   }

   return item.fRecurrence || false;
 }

 /**
  * Get item web URL
  */
 private getItemWebUrl(item: any, source: ISharePointListSource): string {
   return `${source.siteUrl}/Lists/${encodeURIComponent(source.title)}/DispForm.aspx?ID=${item.ID}`;
 }

 /**
  * Get approval status
  */
 private getApprovalStatus(item: any): 'Approved' | 'Pending' | 'Rejected' | 'Draft' {
   const status = item.Status || item.ApprovalStatus || item._ModerationStatus;
   
   if (typeof status === 'string') {
     switch (status.toLowerCase()) {
       case 'approved': return 'Approved';
       case 'pending': return 'Pending';
       case 'rejected': 
       case 'denied': return 'Rejected';
       default: return 'Draft';
     }
   }
   
   return 'Draft';
 }

 /**
  * Extract custom fields
  */
 private extractCustomFields(item: any, source: ISharePointListSource): { [fieldName: string]: unknown } {
   const customFields: { [fieldName: string]: unknown } = {};
   
   source.customFields?.forEach(field => {
     if (item[field.internalName] !== undefined) {
       let value = item[field.internalName];
       
       // Handle complex field types
       if (field.fieldType === 'User' && typeof value === 'object' && value !== null) {
         value = {
           title: value.Title,
           email: value.EMail,
           id: value.ID
         };
       } else if (field.fieldType === 'Lookup' && typeof value === 'object' && value !== null) {
         value = {
           lookupValue: value.LookupValue,
           lookupId: value.LookupId
         };
       }
       
       customFields[field.internalName] = value;
     }
   });
   
   return customFields;
 }

 /**
  * Map attachments
  */
 private mapAttachments(attachments: any[]): ISharePointAttachment[] {
   return attachments.map(att => ({
     fileName: att.FileName,
     serverRelativeUrl: att.ServerRelativeUrl,
     size: att.Length || 0,
     created: new Date(att.TimeCreated || Date.now()),
     createdBy: att.Author?.Title || 'Unknown',
     lastModified: new Date(att.TimeLastModified || Date.now()),
     contentType: att.ContentType || 'application/octet-stream'
   }));
 }

 /**
  * Enrich events with additional data (versions, comments)
  */
 private async enrichEventsWithAdditionalData(events: ISharePointEvent[], source: ISharePointListSource): Promise<void> {
   if (!this.configuration.enableVersionHistory && !this.configuration.enableComments) {
     return;
   }

   const batchSize = 10; // Process in batches to avoid overwhelming the API
   
   for (let i = 0; i < events.length; i += batchSize) {
     const batch = events.slice(i, i + batchSize);
     const promises: Promise<void>[] = [];
     
     for (const event of batch) {
       if (this.configuration.enableVersionHistory) {
         promises.push(this.getEventVersions(event, source));
       }
       
       if (this.configuration.enableComments) {
         promises.push(this.getEventComments(event, source));
       }
     }
     
     try {
       await Promise.all(promises);
     } catch (error) {
       console.warn('Error enriching events with additional data:', error);
     }
   }
 }

 /**
  * Get event versions
  */
 private async getEventVersions(event: ISharePointEvent, source: ISharePointListSource): Promise<void> {
   try {
     const response = await this.context.spHttpClient.get(
       `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items(${event.listItemId})/versions`,
       SPHttpClient.configurations.v1
     );

     if (response.ok) {
       const data = await response.json();
       const versions = data.d?.results || data.value || [];
       
       event.versions = versions.map((version: any) => ({
         versionId: version.ID,
         versionLabel: version.VersionLabel,
         created: new Date(version.Created),
         createdBy: version.Editor?.Title || 'Unknown',
         isCurrentVersion: version.IsCurrentVersion || false
       }));
     }
   } catch (error) {
     console.warn(`Error getting versions for event ${event.id}:`, error);
   }
 }

 /**
  * Get event comments
  */
 private async getEventComments(event: ISharePointEvent, source: ISharePointListSource): Promise<void> {
   try {
     // Note: Comments API may vary based on SharePoint version and configuration
     const response = await this.context.spHttpClient.get(
       `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items(${event.listItemId})/comments`,
       SPHttpClient.configurations.v1
     );

     if (response.ok) {
       const data = await response.json();
       const comments = data.d?.results || data.value || [];
       
       event.comments = comments.map((comment: any) => ({
         id: comment.ID,
         text: comment.Text,
         author: comment.Author?.Title || 'Unknown',
         authorEmail: comment.Author?.EMail,
         created: new Date(comment.Created),
         likeCount: comment.LikeCount || 0,
         replies: comment.Replies?.results || []
       }));
     }
   } catch (error) {
     // Comments might not be available in all SharePoint configurations
     console.debug(`Comments not available for event ${event.id}`);
   }
 }

 /**
  * Search events across SharePoint lists
  */
 public async searchEvents(sources: ISharePointListSource[], query: string, maxResults: number = 50): Promise<ISharePointEvent[]> {
   const allEvents: ISharePointEvent[] = [];
   const searchTerms = query.toLowerCase().split(' ').filter(term => term.length > 0);

   if (searchTerms.length === 0) {
     return allEvents;
   }

   console.log(`Searching for "${query}" across ${sources.length} lists`);

   for (const source of sources.filter(s => s.isEnabled && s.permissions.canRead)) {
     try {
       // Build search filter for SharePoint
       const searchFilter = this.buildSearchFilter(searchTerms, source);
       
       const selectFields = this.buildSelectFields(source);
       
       const apiUrl = `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items?` +
         `$select=${selectFields}&` +
         `$expand=AttachmentFiles,Author,Editor,AssignedTo&` +
         `$filter=${searchFilter}&` +
         `$orderby=${this.getOrderByField(source)}&` +
         `$top=${maxResults}`;

       const response: SPHttpClientResponse = await this.context.spHttpClient.get(
         apiUrl,
         SPHttpClient.configurations.v1
       );

       if (response.ok) {
         const data = await response.json();
         const items = data.d?.results || data.value || [];

         for (const item of items) {
           const event = this.mapListItemToEvent(item, source);
           if (event) {
             allEvents.push(event);
           }
         }
       }
     } catch (error) {
       console.error(`Error searching events in ${source.title}:`, error);
     }
   }

   console.log(`Found ${allEvents.length} events matching search query`);
   
   return allEvents
     .sort((a, b) => a.start.getTime() - b.start.getTime())
     .slice(0, maxResults);
 }

 /**
  * Build search filter for SharePoint list
  */
 private buildSearchFilter(searchTerms: string[], source: ISharePointListSource): string {
   const searchableFields = ['Title'];
   
   // Add description field if available
   if (source.fieldMappings?.descriptionField) {
     searchableFields.push(source.fieldMappings.descriptionField);
   } else {
     searchableFields.push('Description', 'Body', 'Comments');
   }
   
   // Add location field if available
   if (source.fieldMappings?.locationField) {
     searchableFields.push(source.fieldMappings.locationField);
   } else {
     searchableFields.push('Location');
   }
   
   // Add category field if available
   if (source.fieldMappings?.categoryField) {
     searchableFields.push(source.fieldMappings.categoryField);
   } else {
     searchableFields.push('Category');
   }

   const termFilters = searchTerms.map(term => {
     const fieldFilters = searchableFields.map(field => 
       `substringof('${term}',${field})`
     );
     return `(${fieldFilters.join(' or ')})`;
   });

   return termFilters.join(' and ');
 }

 /**
  * Get events for specific date range
  */
 public async getEventsForDateRange(
   sources: ISharePointListSource[], 
   startDate: Date, 
   endDate: Date, 
   maxEvents: number = 1000
 ): Promise<ISharePointEvent[]> {
   const allEvents: ISharePointEvent[] = [];

   console.log(`Getting events for date range: ${DateUtils.formatDate(startDate)} to ${DateUtils.formatDate(endDate)}`);

   for (const source of sources.filter(s => s.isEnabled && s.permissions.canRead)) {
     try {
       const dateFilter = this.buildDateRangeFilter(source, startDate, endDate);
       const selectFields = this.buildSelectFields(source);
       const orderBy = this.getOrderByField(source);

       const apiUrl = `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items?` +
         `$select=${selectFields}&` +
         `$expand=AttachmentFiles,Author,Editor,AssignedTo&` +
         `$filter=${dateFilter}&` +
         `$orderby=${orderBy}&` +
         `$top=${maxEvents}`;

       const response: SPHttpClientResponse = await this.context.spHttpClient.get(
         apiUrl,
         SPHttpClient.configurations.v1
       );

       if (response.ok) {
         const data = await response.json();
         const items = data.d?.results || data.value || [];

         for (const item of items) {
           const event = this.mapListItemToEvent(item, source);
           if (event && event.start >= startDate && event.start <= endDate) {
             allEvents.push(event);
           }
         }
       }
     } catch (error) {
       console.error(`Error getting events from ${source.title}:`, error);
     }
   }

   console.log(`Retrieved ${allEvents.length} events for date range`);
   
   return allEvents
     .sort((a, b) => a.start.getTime() - b.start.getTime())
     .slice(0, maxEvents);
 }

 /**
  * Build date range filter
  */
 private buildDateRangeFilter(source: ISharePointListSource, startDate: Date, endDate: Date): string {
   const startDateField = source.fieldMappings?.startDateField || this.getDefaultStartDateField(source);
   
   return `${startDateField} ge datetime'${startDate.toISOString()}' and ${startDateField} le datetime'${endDate.toISOString()}'`;
 }

 /**
  * Update list configuration
  */
 public updateConfiguration(config: Partial<ISharePointListConfiguration>): void {
   this.configuration = { ...this.configuration, ...config };
 }

 /**
  * Get current configuration
  */
 public getConfiguration(): ISharePointListConfiguration {
   return { ...this.configuration };
 }

 /**
  * Check calendar permissions
  */
 public async checkCalendarPermissions(source: ISharePointListSource): Promise<{ canRead: boolean; canWrite: boolean }> {
   try {
     const permissions = await this.getListPermissions(source.id, source.siteUrl);
     return {
       canRead: permissions.canRead,
       canWrite: permissions.canWrite
     };
   } catch (error) {
     console.warn(`Could not check permissions for ${source.title}:`, error);
     return { canRead: false, canWrite: false };
   }
 }

 /**
  * Get calendar statistics
  */
 public async getCalendarStatistics(source: ISharePointListSource): Promise<{ totalItems: number; recentItems: number; upcomingItems: number }> {
   try {
     const now = new Date();
     const oneWeekAgo = DateUtils.subtractTime(now, 7, 'days');
     const oneMonthFromNow = DateUtils.addTime(now, 30, 'days');

     // Get total count
     const totalResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
       `${source.siteUrl}/_api/web/lists(guid'${source.id}')/ItemCount`,
       SPHttpClient.configurations.v1
     );

     // Get recent items count
     const recentFilter = `Created ge datetime'${oneWeekAgo.toISOString()}'`;
     const recentResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
       `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items?$select=Id&$filter=${recentFilter}&$top=1000`,
       SPHttpClient.configurations.v1
     );

     // Get upcoming events count
     const startDateField = source.fieldMappings?.startDateField || this.getDefaultStartDateField(source);
     const upcomingFilter = `${startDateField} ge datetime'${now.toISOString()}' and ${startDateField} le datetime'${oneMonthFromNow.toISOString()}'`;
     const upcomingResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
       `${source.siteUrl}/_api/web/lists(guid'${source.id}')/items?$select=Id&$filter=${upcomingFilter}&$top=1000`,
       SPHttpClient.configurations.v1
     );

     let totalItems = 0;
     let recentItems = 0;
     let upcomingItems = 0;

     if (totalResponse.ok) {
       const totalData = await totalResponse.json();
       totalItems = totalData.d || totalData.value || 0;
     }

     if (recentResponse.ok) {
       const recentData = await recentResponse.json();
       recentItems = (recentData.d?.results || recentData.value || []).length;
     }

     if (upcomingResponse.ok) {
       const upcomingData = await upcomingResponse.json();
       upcomingItems = (upcomingData.d?.results || upcomingData.value || []).length;
     }

     return { totalItems, recentItems, upcomingItems };
   } catch (error) {
     console.warn(`Could not get statistics for ${source.title}:`, error);
     return { totalItems: 0, recentItems: 0, upcomingItems: 0 };
   }
 }
}