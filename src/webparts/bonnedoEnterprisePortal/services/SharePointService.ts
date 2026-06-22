import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';
import {
  mapInventoryMovementRecord,
  mapInventoryRecord,
  mapMaterialMasterRecord,
  mapWarehouseRecord,
} from '../utils/materialDataMappers';
import {
  IInventoryMovementRecord,
  IInventoryRecord,
  IMaterialMasterRecord,
  IWarehouseRecord,
} from '../models/DataModels';

export interface IListItem {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
  ID: number;
}

export interface ISharePointListData {
  value: IListItem[];
}

export class SharePointService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;
  private listFieldCache: Map<string, string[]> = new Map();

  private readonly fieldNameAliases: Record<string, string[]> = {
    From_Location: ['From_Location', 'From_x0020_Location'],
    To_Location: ['To_Location', 'To_x0020_Location'],
    Project_Code: ['Project_Code', 'Project_x0020_Code'],
    Material_Code: ['Material_Code', 'Material_x0020_Code'],
    GRN_Number: ['GRN_Number', 'GRN_x0020_Number'],
    PO_Number: ['PO_Number', 'PO_x0020_Number'],
    QRCodeURL: ['QRCodeURL', 'qrcodeurl', 'QR_Code', 'QRCode'],
  };

  constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
    this.spHttpClient = spHttpClient;
    this.pageContext = pageContext;
  }

  private async getListFieldInternalNames(listName: string): Promise<string[]> {
    if (this.listFieldCache.has(listName)) {
      return this.listFieldCache.get(listName) || [];
    }

    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/fields?$select=InternalName,Hidden&$filter=Hidden eq false`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return [];
    }

    const data: ISharePointListData = await response.json();
    const fieldNames = (data.value || []).map((field) => String(field.InternalName || '')).filter(Boolean);
    this.listFieldCache.set(listName, fieldNames);
    return fieldNames;
  }

  private resolveFieldName(fieldNames: string[], fieldName: string): string | undefined {
    if (fieldNames.includes(fieldName)) {
      return fieldName;
    }

    const aliases = this.fieldNameAliases[fieldName] || [];
    return aliases.find((alias) => fieldNames.includes(alias));
  }

  private async sanitizeItemDataForList(
    listName: string,
    itemData: { [key: string]: any }
  ): Promise<{ [key: string]: any }> {
    const fieldNames = await this.getListFieldInternalNames(listName);

    if (fieldNames.length === 0) {
      return itemData;
    }

    const sanitized: { [key: string]: any } = {};
    const removedFields: string[] = [];

    Object.keys(itemData).forEach((fieldName) => {
      const resolvedFieldName = this.resolveFieldName(fieldNames, fieldName);

      if (!resolvedFieldName) {
        removedFields.push(fieldName);
        return;
      }

      sanitized[resolvedFieldName] = itemData[fieldName];
    });

    if (removedFields.length > 0) {
      console.warn(`SharePoint list '${listName}' does not contain these fields: ${removedFields.join(', ')}`);
    }

    return sanitized;
  }

  /**
   * Fetch data from a SharePoint list with optional filtering
   * @param listName - Name of the SharePoint list
   * @param filterQuery - OData filter query (optional)
   * @param top - Number of items to retrieve (default: 100)
   * @param skip - Number of items to skip for pagination (default: 0)
   * @returns Promise with list items
   */
  public async getListData(
    listName: string,
    filterQuery?: string,
    top: number = 100,
    skip: number = 0,
    expand?: string
  ): Promise<IListItem[]> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      let url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=${top}&$skip=${skip}`;

      // Add filter if provided
      if (filterQuery) {
        // Encode the filter query to handle special characters properly
        url += `&$filter=${encodeURIComponent(filterQuery)}`;
      }

      // Add expand if provided
      if (expand) {
        url += `&$expand=${expand}`;
        // Build proper $select for multiple expanded fields
        const expandFields = expand.split(',').map(f => f.trim());
        const selectParts = expandFields.map(f => `${f}/ID,${f}/Title`).join(',');
        url += `&$select=*,${selectParts}`;
      } else {
        // Add select to get all fields
        url += '&$select=*';
      }

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        // Log more details for debugging
        const errorText = await response.text();
        console.error('SharePoint API Error:', {
          status: response.status,
          statusText: response.statusText,
          url: url,
          error: errorText
        });
        throw new Error(`Failed to fetch list data: ${response.statusText} - ${errorText}`);
      }

      const data: ISharePointListData = await response.json();
      return data.value;
    } catch (error) {
      console.error('Error fetching SharePoint list data:', error);
      throw error;
    }
  }

  /**
   * Get total count of items in a list
   * @param listName - Name of the SharePoint list
   * @param filterQuery - OData filter query (optional)
   * @returns Promise with item count
   */
  public async getListItemCount(
    listName: string,
    filterQuery?: string
  ): Promise<number> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      let url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=0`;

      if (filterQuery) {
        // Encode the filter query to handle special characters properly
        url += `&$filter=${encodeURIComponent(filterQuery)}`;
      }

      // Use $inlinecount for better compatibility
      url += '&$inlinecount=allpages';

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch list count: ${response.statusText}`);
      }

      const data = await response.json();
      // OData v3/v4 count properties
      const count = data['odata.count'] || data['__count'] || (data.value ? data.value.length : 0);
      return typeof count === 'string' ? parseInt(count, 10) : count;
    } catch (error) {
      console.error('Error fetching SharePoint list count:', error);
      throw error;
    }
  }

  /**
   * Get the latest item from a list based on creation date
   * @param listName - Name of the SharePoint list
   * @returns Promise with the latest item or undefined
   */
  public async getLatestItem(listName: string): Promise<IListItem | undefined> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      // Only select ID and Title to avoid 400 errors with non-existent fields
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=1&$orderby=Created desc&$select=ID,Title`;

      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return undefined;

      const data = await response.json();
      return data.value && data.value.length > 0 ? data.value[0] : undefined;
    } catch (error) {
      console.error('Error fetching latest item:', error);
      return undefined;
    }
  }

  /**
   * Get items sorted by a specific field to find the latest based on that field
   * @param listName - Name of the SharePoint list
   * @param sortField - Field to sort by
   * @param selectFields - Fields to select
   * @param top - Number of items to retrieve
   * @returns Promise with sorted items
   */
  public async getItemsSorted(
    listName: string,
    sortField: string,
    selectFields: string = 'ID,Title',
    top: number = 1,
    orderDescending: boolean = true
  ): Promise<IListItem[]> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const orderDir = orderDescending ? 'desc' : 'asc';
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=${top}&$orderby=${sortField} ${orderDir}&$select=${selectFields}`;

      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return [];

      const data = await response.json();
      return data.value || [];
    } catch (error) {
      console.error('Error fetching sorted items:', error);
      return [];
    }
  }

  /**
   * Get list fields/columns
   * @param listName - Name of the SharePoint list
   * @returns Promise with list fields
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async getListFields(listName: string): Promise<any[]> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/fields?$filter=Hidden eq false`;

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch list fields: ${response.statusText}`);
      }

      const data = await response.json();
      return data.value;
    } catch (error) {
      console.error('Error fetching SharePoint list fields:', error);
      throw error;
    }
  }

  /**
   * Get choices for a choice field
   * @param listName - Name of the SharePoint list
   * @param fieldName - Name of the field
   * @returns Promise with array of choices
   */
  public async getFieldChoices(listName: string, fieldName: string): Promise<string[]> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/fields/getByTitle('${fieldName}')`;

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch field: ${response.statusText}`);
      }

      const field = await response.json();
      return field.Choices || [];
    } catch (error) {
      console.error('Error fetching field choices:', error);
      throw error;
    }
  }

  /**
   * Add a new choice to a choice field
   * @param listName - Name of the SharePoint list
   * @param fieldName - Name of the field
   * @param newChoice - The new choice to add
   * @returns Promise
   */
  public async addFieldChoice(listName: string, fieldName: string, newChoice: string): Promise<void> {
    try {
      // First get current choices
      const currentChoices = await this.getFieldChoices(listName, fieldName);

      // Check if choice already exists
      if (currentChoices.includes(newChoice)) {
        return; // Already exists
      }

      // Add new choice
      const updatedChoices = [...currentChoices, newChoice];

      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/fields/getByTitle('${fieldName}')`;

      const updateData = {
        Choices: updatedChoices,
      };

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*',
          },
          body: JSON.stringify(updateData),
        }
      );

      if (!response.ok) {
        console.warn(`Failed to add choice "${newChoice}" to field ${fieldName}. User may not have permissions.`);
        // Don't throw error, just log warning - the item will still be saved
      }
    } catch (error) {
      console.warn('Error adding field choice:', error);
      // Don't throw error - the item will still be saved
    }
  }

  /**
   * Update a list item
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to update
   * @param updates - Object containing fields to update
   * @returns Promise with updated item
   */
  public async updateListItem(
    listName: string,
    itemId: number,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    updates: { [key: string]: any }
  ): Promise<void> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})`;

      const bodyData = await this.sanitizeItemDataForList(listName, updates);
      if (Object.keys(bodyData).length === 0) {
        return;
      }

      const body = JSON.stringify(bodyData);

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*',
          },
          body: body,
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to update list item: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error updating SharePoint list item:', error);
      throw error;
    }
  }

  /**
   * Create a new list item
   * @param listName - Name of the SharePoint list
   * @param itemData - Object containing fields to create
   * @returns Promise with created item
   */
  public async createListItem(
    listName: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    itemData: { [key: string]: any }
  ): Promise<IListItem> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items`;

      const sanitizedItemData = await this.sanitizeItemDataForList(listName, itemData);
      if (Object.keys(sanitizedItemData).length === 0) {
        throw new Error(`No valid fields were found for SharePoint list '${listName}'.`);
      }

      const body = JSON.stringify(sanitizedItemData);

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
          body: body,
        }
      );

      if (!response.ok) {
        // Log more details for debugging
        const errorText = await response.text();
        console.error('SharePoint API Error:', {
          status: response.status,
          statusText: response.statusText,
          url: url,
          body: body,
          error: errorText
        });
        throw new Error(`Failed to create list item: ${response.statusText} - ${errorText}`);
      }

      const createdItem: IListItem = await response.json();
      return createdItem;
    } catch (error) {
      console.error('Error creating SharePoint list item:', error);
      throw error;
    }
  }

  /**
   * Append text to a list item field
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item
   * @param fieldName - Name of the field to append to
   * @param textToAppend - Text to append
   * @returns Promise with updated item
   */
  public async appendToField(
    listName: string,
    itemId: number,
    fieldName: string,
    textToAppend: string
  ): Promise<void> {
    try {
      // First, get the current value
      const webUrl = this.pageContext.web.absoluteUrl;
      const getUrl = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})?$select=${fieldName}`;

      const getResponse: SPHttpClientResponse = await this.spHttpClient.get(
        getUrl,
        SPHttpClient.configurations.v1
      );

      if (!getResponse.ok) {
        throw new Error(`Failed to fetch field value: ${getResponse.statusText}`);
      }

      const item = await getResponse.json();
      const currentValue = item[fieldName] || '';
      const newValue = currentValue ? `${currentValue}\n${textToAppend}` : textToAppend;

      // Update with appended value
      await this.updateListItem(listName, itemId, { [fieldName]: newValue });
    } catch (error) {
      console.error('Error appending to SharePoint list field:', error);
      throw error;
    }
  }

  // ========== LOOKUP HELPER METHODS ==========

  /**
   * Get projects from ENT_Project_Master list for dropdown
   */
  public async getProjects(): Promise<IListItem[]> {
    return this.getListData('ENT_Project_Master', undefined, 100);
  }

  /**
   * Get materials from ENT_Materials_Master list for dropdown
   */
  public async getMaterials(): Promise<IListItem[]> {
    return this.getListData(SHAREPOINT_LISTS.MATERIALS_MASTER, undefined, 100);
  }

  /**
   * Get material master records for the Material Management module
   */
  public async getMaterialMasterRecords(top: number = 500): Promise<IMaterialMasterRecord[]> {
    const items = await this.getOrderedListData(SHAREPOINT_LISTS.MATERIALS_MASTER, top, 'Material_Code desc');
    return items.map(mapMaterialMasterRecord);
  }

  /**
   * Get warehouses from ENT_Warehouses_Master
   */
  public async getWarehouses(top: number = 50): Promise<IWarehouseRecord[]> {
    const items = await this.getOrderedListData(SHAREPOINT_LISTS.WAREHOUSES_MASTER, top, 'Title asc');
    return items.map(mapWarehouseRecord);
  }

  /**
   * Get inventory records from ENT_Inventory_Register
   */
  public async getInventoryRecords(top: number = 1000): Promise<IInventoryRecord[]> {
    const items = await this.getOrderedListData(SHAREPOINT_LISTS.INVENTORY_REGISTER, top, 'Material_Code asc');
    return items.map(mapInventoryRecord);
  }

  /**
   * Get inventory movement records from ENT_Inventory_Movements_Register
   */
  public async getInventoryMovements(top: number = 500): Promise<IInventoryMovementRecord[]> {
    const items = await this.getOrderedListData(SHAREPOINT_LISTS.INVENTORY_MOVEMENTS_REGISTER, top, 'Created desc');
    return items.map(mapInventoryMovementRecord);
  }

  /**
   * Build SharePoint native new-item form URL for a list
   */
  public getListNewFormUrl(listName: string): string {
    const encodedListName = listName.replace(/ /g, '%20');
    return `${this.pageContext.web.absoluteUrl}/Lists/${encodedListName}/NewForm.aspx`;
  }

  /**
   * Fetch list items with OData $orderby
   */
  private async getOrderedListData(
    listName: string,
    top: number,
    orderBy: string,
    filterQuery?: string
  ): Promise<IListItem[]> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      let url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=${top}&$orderby=${encodeURIComponent(orderBy)}&$select=*`;

      if (filterQuery) {
        url += `&$filter=${encodeURIComponent(filterQuery)}`;
      }

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to fetch list data: ${response.statusText} - ${errorText}`);
      }

      const data: ISharePointListData = await response.json();
      return data.value;
    } catch (error) {
      console.error('Error fetching ordered SharePoint list data:', error);
      throw error;
    }
  }

  /**
   * Get vendors from ENT_Vendors_Master list for dropdown
   */
  public async getVendors(): Promise<IListItem[]> {
    return this.getListData('ENT_Vendors_Master', undefined, 100);
  }

  /**
   * Get material requests from PRC_Material_Request_Register for dropdown
   */
  public async getMaterialRequests(): Promise<IListItem[]> {
    return this.getListData('PRC_Material_Request_Register', undefined, 100);
  }

  /**
   * Get the current logged-in user from SharePoint
   */
  public async getCurrentUser(): Promise<IListItem | undefined> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/currentUser`;

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to get current user:', response.statusText);
        return undefined;
      }

      const data = await response.json();
      return {
        ID: data.Id,
        Title: data.Title,
        Email: data.Email,
      };
    } catch (error) {
      console.error('Error getting current user:', error);
      return undefined;
    }
  }

  /**
   * Get purchase requisitions from PRC_Purchase_Requisition_Register for dropdown
   */
  public async getPurchaseRequisitions(): Promise<IListItem[]> {
    return this.getListData('PRC_Purchase_Requisition_Register', undefined, 100);
  }

  /**
   * Get purchase orders from PRC_Purchase_Order_Register for dropdown
   */
  public async getPurchaseOrders(): Promise<IListItem[]> {
    return this.getListData('PRC_Purchase_Order_Register', undefined, 100);
  }

  // ========== FINANCE METHODS ==========

  /**
   * Get payment requests from FIN_Payment_Request_Register
   */
  public async getPaymentRequests(filterQuery?: string): Promise<IListItem[]> {
    return this.getListData('FIN_Payment_Request_Register', filterQuery, 100);
  }

  /**
   * Get approved payments (filter by Approval_Status = Approved)
   */
  public async getApprovedPayments(): Promise<IListItem[]> {
    return this.getListData('FIN_Payment_Request_Register', "Approval_Status eq 'Approved'", 100);
  }

  /**
   * Get expenses from FIN_Expense_Register
   */
  public async getExpenses(): Promise<IListItem[]> {
    return this.getListData('FIN_Expense_Register', undefined, 100);
  }

  // ========== SETTINGS METHODS ==========

  /**
   * Get approval matrix from SYS_Approval_Matrix
   */
  public async getApprovalMatrix(): Promise<IListItem[]> {
    return this.getListData('SYS_Approval_Matrix', undefined, 100);
  }

  /**
   * Get system configuration from SYS_System_Config
   */
  public async getSystemConfig(): Promise<IListItem[]> {
    return this.getListData('SYS_System_Config', undefined, 10);
  }

  /**
   * Get user roles from SYS_User_Roles
   * Expands the User field to properly retrieve user information
   */
  public async getUserRoles(): Promise<IListItem[]> {
    return this.getListData('SYS_User_Roles', undefined, 100, 0, 'User');
  }

  /**
   * Save or update a user role in SYS_User_Roles list
   * If user already exists, updates their role; otherwise creates new entry
   * @param userId - The SharePoint user ID
   * @param userTitle - The user's display name
   * @param role - The role to assign (Admin, Viewer, etc.)
   * @param department - Optional department
   * @param permissions - Optional permissions description
   * @returns Promise with created/updated item
   */
  public async saveUserRole(
    userId: number,
    userTitle: string,
    role: string,
    department?: string,
    permissions?: string
  ): Promise<IListItem> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;

      // Check if user already exists in SYS_User_Roles
      const existingRoles = await this.getListData(
        'SYS_User_Roles',
        `User/ID eq ${userId}`,
        1,
        0,
        'User'
      );

      if (existingRoles && existingRoles.length > 0) {
        // User exists, update their role
        const existingItem = existingRoles[0];
        const updateUrl = `${webUrl}/_api/web/lists/getByTitle('SYS_User_Roles')/items(${existingItem.ID})`;

        const updateData: { [key: string]: any } = {
          'Role': role
        };

        if (department) {
          updateData['Department'] = department;
        }
        if (permissions) {
          updateData['Permissions'] = permissions;
        }

        const response = await this.spHttpClient.post(
          updateUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json',
              'Content-Type': 'application/json',
              'X-HTTP-Method': 'MERGE',
              'If-Match': '*'
            },
            body: JSON.stringify(updateData)
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Failed to update user role:', errorText);
          throw new Error(`Failed to update user role: ${response.statusText}`);
        }

        console.log(`Updated role for user ${userId} to ${role}`);
        return { ...existingItem, Role: role };
      } else {
        // User doesn't exist, create new entry
        // Use simplest format without metadata - let SharePoint infer the type
        const newItemData: { [key: string]: any } = {
          'Title': userTitle,
          'Role': role,
          'UserId': userId
        };

        if (department) {
          newItemData['Department'] = department;
        }
        if (permissions) {
          newItemData['Permissions'] = permissions;
        }

        // Direct POST without metadata
        const url = `${webUrl}/_api/web/lists/getByTitle('SYS_User_Roles')/items`;
        const response = await this.spHttpClient.post(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json',
              'Content-Type': 'application/json',
            },
            body: JSON.stringify(newItemData)
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Failed to create user role:', errorText);
          throw new Error(`Failed to create user role: ${response.statusText}`);
        }

        const createdItem = await response.json();
        console.log(`Created new user role entry for ${userTitle} with role ${role}`);
        return createdItem;
      }
    } catch (error) {
      console.error('Error saving user role:', error);
      throw error;
    }
  }

  /**
   * Get notification settings from SYS_Notification_Settings
   */
  public async getNotificationSettings(): Promise<IListItem[]> {
    return this.getListData('SYS_Notification_Settings', undefined, 10);
  }

  // ========== REPORTS METHODS ==========

  /**
   * Delete a list item
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to delete
   * @returns Promise
   */
  public async deleteListItem(listName: string, itemId: number): Promise<void> {
    try {
      const webUrl = this.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})`;

      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*',
          },
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to delete list item: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error deleting SharePoint list item:', error);
      throw error;
    }
  }

  /**
   * Get data for reports
   */
  public async getReportsData(listName: string, filterQuery?: string): Promise<IListItem[]> {
    return this.getListData(listName, filterQuery, 500);
  }

  /**
   * Initialize Material Request approval workflow on creation (STEP 1 complete)
   */
  public async initializeMaterialRequestApproval(itemId: number, module: string = 'Material Request'): Promise<void> {
    try {
      const now = new Date().toISOString();

      let currentApprover: string | null = null;
      try {
        const approvalMatrix = await this.getListData(
          'SYS_Approval_Matrix',
          `Module eq '${module}' and Stage eq 1`,
          1,
          0,
          'Approver_User'
        );
        if (approvalMatrix.length > 0) {
          const approverField = approvalMatrix[0].Approver_User || approvalMatrix[0].Approver || approvalMatrix[0].Title;
          currentApprover = typeof approverField === 'object' ? (approverField.Title || approverField.Email || null) : (approverField ? String(approverField) : null);
        }
      } catch (e) { console.warn('Approval Matrix query failed'); }

      const initialHistory = JSON.stringify([{
        stage: 1,
        approver: currentApprover || '',
        action: 'Pending',
        date: now,
        comment: ''
      }]);

      const updateData: any = {
        Approval_Status: 'Pending',
        Approval_Level: 1,
        Current_Approver: currentApprover,
        Approval_Started_On: now,
        Approval_History: initialHistory,
        Status: 'Pending Approval'
      };

      await this.updateListItem('PRC_Material_Request_Register', itemId, updateData);
    } catch (error) {
      console.error('Error initializing MR approval:', error);
    }
  }

  /**
   * Process approve/reject action for Material Request (STEP 3 + 4)
   */
  public async processMaterialRequestApprovalAction(
    itemId: number,
    action: 'approve' | 'reject',
    approverEmail: string,
    comment: string = ''
  ): Promise<void> {
    const now = new Date().toISOString();
    const items = await this.getListData('PRC_Material_Request_Register', `ID eq ${itemId}`, 1);
    if (!items.length) return;
    const item = items[0];

    let history: any[] = [];
    try { history = JSON.parse(item.Approval_History || '[]'); } catch {}

    if (action === 'reject') {
      history.push({ stage: item.Approval_Level || 1, approver: approverEmail, action: 'Rejected', date: now, comment });
      await this.updateListItem('PRC_Material_Request_Register', itemId, {
        Approval_Status: 'Rejected',
        Status: 'Rejected',
        Approval_History: JSON.stringify(history)
      });
      return;
    }

    // Approve path - update last pending entry or add new Approved entry
    const currentLevel = item.Approval_Level || 1;
    const lastEntry = history[history.length - 1];
    if (lastEntry && lastEntry.action === 'Pending') {
      lastEntry.action = 'Approved';
      lastEntry.approver = approverEmail;
      lastEntry.date = now;
      lastEntry.comment = comment;
    } else {
      history.push({ stage: currentLevel, approver: approverEmail, action: 'Approved', date: now, comment });
    }

    // Fetch next stage from matrix
    const nextStage = currentLevel + 1;
    let nextApprover: string | null = null;
    try {
      const matrix = await this.getListData('SYS_Approval_Matrix', `Module eq 'Material Request' and Stage eq ${nextStage}`, 1, 0, 'Approver_User');
      if (matrix.length > 0) {
        const f = matrix[0].Approver_User || matrix[0].Approver || matrix[0].Title;
        nextApprover = typeof f === 'object' ? (f.Title || f.Email || null) : (f ? String(f) : null);
      }
    } catch {}

    if (nextApprover) {
      await this.updateListItem('PRC_Material_Request_Register', itemId, {
        Approval_Level: nextStage,
        Current_Approver: nextApprover,
        Approval_History: JSON.stringify(history)
      });
    } else {
      // Last stage - finalize
      await this.updateListItem('PRC_Material_Request_Register', itemId, {
        Approval_Status: 'Approved',
        Approval_Completed_On: now,
        Status: 'Approved',
        Current_Approver: null,
        Approval_History: JSON.stringify(history)
      });
    }
  }

  /**
   * Initialize Purchase Requisition approval workflow
   */
  public async initializePurchaseRequisitionApproval(itemId: number): Promise<void> {
    try {
      const now = new Date().toISOString();

      let currentApprover: string | null = null;
      try {
        const approvalMatrix = await this.getListData(
          'SYS_Approval_Matrix',
          `Module eq 'Purchase Requisition' and Stage eq 1`,
          1,
          0,
          'Approver_User'
        );
        if (approvalMatrix.length > 0) {
          const approverField = approvalMatrix[0].Approver_User || approvalMatrix[0].Approver || approvalMatrix[0].Title;
          currentApprover = typeof approverField === 'object' ? (approverField.Title || approverField.Email || null) : (approverField ? String(approverField) : null);
        }
      } catch (e) { console.warn('Approval Matrix query failed for PR'); }

      const initialHistory = JSON.stringify([{
        stage: 1,
        approver: currentApprover || '',
        action: 'Pending',
        date: now,
        comment: ''
      }]);

      const updateData: any = {
        Approval_Status: 'Pending',
        Approval_Level: 1,
        Current_Approver: currentApprover,
        Approval_Started_On: now,
        Approval_History: initialHistory,
        Status: 'Pending Approval'
      };

      await this.updateListItem('PRC_Purchase_Requisition_Register', itemId, updateData);
    } catch (error) {
      console.error('Error initializing PR approval:', error);
    }
  }

  /**
   * Process approve/reject action for Purchase Requisition (two-stage)
   */
  public async processPurchaseRequisitionApprovalAction(
    itemId: number,
    action: 'approve' | 'reject',
    approverEmail: string,
    comment: string = ''
  ): Promise<void> {
    const now = new Date().toISOString();
    const items = await this.getListData('PRC_Purchase_Requisition_Register', `ID eq ${itemId}`, 1);
    if (!items.length) return;
    const item = items[0];

    let history: any[] = [];
    try { history = JSON.parse(item.Approval_History || '[]'); } catch {}

    if (action === 'reject') {
      history.push({ stage: item.Approval_Level || 1, approver: approverEmail, action: 'Rejected', date: now, comment });
      await this.updateListItem('PRC_Purchase_Requisition_Register', itemId, {
        Approval_Status: 'Rejected',
        Status: 'Rejected',
        Approval_History: JSON.stringify(history)
      });
      return;
    }

    const currentLevel = item.Approval_Level || 1;
    const lastEntry = history[history.length - 1];
    if (lastEntry && lastEntry.action === 'Pending') {
      lastEntry.action = 'Approved';
      lastEntry.approver = approverEmail;
      lastEntry.date = now;
      lastEntry.comment = comment;
    } else {
      history.push({ stage: currentLevel, approver: approverEmail, action: 'Approved', date: now, comment });
    }

    const nextStage = currentLevel + 1;
    let nextApprover: string | null = null;
    try {
      const matrix = await this.getListData('SYS_Approval_Matrix', `Module eq 'Purchase Requisition' and Stage eq ${nextStage}`, 1, 0, 'Approver_User');
      if (matrix.length > 0) {
        const f = matrix[0].Approver_User || matrix[0].Approver || matrix[0].Title;
        nextApprover = typeof f === 'object' ? (f.Title || f.Email || null) : (f ? String(f) : null);
      }
    } catch {}

    if (nextApprover) {
      await this.updateListItem('PRC_Purchase_Requisition_Register', itemId, {
        Approval_Level: nextStage,
        Current_Approver: nextApprover,
        Approval_History: JSON.stringify(history)
      });
    } else {
      await this.updateListItem('PRC_Purchase_Requisition_Register', itemId, {
        Approval_Status: 'Approved',
        Approval_Completed_On: now,
        Status: 'Approved',
        Current_Approver: null,
        Approval_History: JSON.stringify(history)
      });
    }
  }

  /**
   * Send escalation reminders for Purchase Requisitions pending > 24 hours
   * This can be called manually or triggered from a dashboard load.
   */
  public async sendPurchaseRequisitionEscalationReminders(): Promise<void> {
    try {
      const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();

      // Find PRs that are still pending and created more than 24 hours ago
      const pendingPRs = await this.getListData(
        'PRC_Purchase_Requisition_Register',
        `Approval_Status eq 'Pending' and Created lt '${oneDayAgo}'`,
        50
      );

      console.log(`[Escalation] Found ${pendingPRs.length} PRs pending for more than 24 hours`);

      // In a real implementation, you would send emails here using NotificationService
      // For now we just log them
      pendingPRs.forEach(pr => {
        console.log(`[Escalation] Reminder needed for PR: ${pr.Title} (ID: ${pr.ID})`);
      });
    } catch (error) {
      console.error('Error sending PR escalation reminders:', error);
    }
  }

  /**
   * Initialize Purchase Order approval with dynamic threshold logic
   */
   /**
    * PO Approval columns required on PRC_Purchase_Order_Register (for the approval workflow to function):
    *   - Status (set to 'Pending Approval')
    *   - Approval_Level (stage number)
    *   - Current_Approver (string or user object)
    *   - Approval_Status ('Pending' / 'Approved')
    *   - Approval_Started_On, Approval_Completed_On
    *   - TotalAmount / Amount (the value passed for threshold evaluation)
    *   - PO_Number / Title
    *
    * On SYS_Approval_Matrix (for 'Purchase Order' module) — exact columns:
    *   - Module (Choice) = 'Purchase Order'
    *   - Stage (Number)
    *   - Min_Amount (Number)
    *   - Max_Amount (Number)
    *   - Approver_Role (Choice) — e.g. "CFO" for stage 2
    *   - Approver_User (Person or Group)
    *
    * Logic implemented below:
    *   - Stage 1 if Total PO Amount is within [Min_Amount, Max_Amount] of stage 1.
    *   - If beyond stage 1's Max_Amount → use stage 2 where Approver_Role contains "CFO".
    */
  public async initializePurchaseOrderApproval(itemId: number, totalAmount: number = 0): Promise<void> {
    try {
      const now = new Date().toISOString();

      // Fetch approval matrix for Purchase Order (ensure we get threshold columns)
      const approvalMatrix = await this.getListData(
        'SYS_Approval_Matrix',
        `Module eq 'Purchase Order'`,
        20,
        0,
        'Approver_User'
      );

      // Map stages using the exact SYS_Approval_Matrix columns
      const stages = approvalMatrix
        .map((item: any) => ({
          stage: item.Stage || 1,
          approver: this.extractApproverName(item.Approver_User || item.Approver),
          thresholdMin: item.Min_Amount ?? 0,
          thresholdMax: item.Max_Amount ?? Infinity,
          role: (item.Approver_Role || '').toString().toLowerCase(),
        }))
        .sort((a, b) => a.stage - b.stage);

      let currentStage = 1;
      let currentApprover: string | null = null;

      // Stage 1: within its min/max threshold
      const stage1 = stages.find(s => s.stage === 1);
      const stage2 = stages.find(s => s.stage === 2);

      if (stage1 && totalAmount >= stage1.thresholdMin && totalAmount <= stage1.thresholdMax) {
        currentStage = 1;
        currentApprover = stage1.approver;
      } else if (stage2 && totalAmount > (stage1?.thresholdMax || 0)) {
        // Beyond stage 1 Max_Amount → stage 2 where Approver_Role is CFO
        const cfoApprover = stages.find(s =>
          s.stage === 2 &&
          (s.role.includes('cfo') || s.approver?.toLowerCase().includes('cfo'))
        );
        if (cfoApprover) {
          currentStage = 2;
          currentApprover = cfoApprover.approver;
        } else if (stage2) {
          // Fallback to stage 2 even if role doesn't perfectly match "CFO"
          currentStage = 2;
          currentApprover = stage2.approver;
        }
      } else {
        // Fallback: pick the highest stage the amount qualifies for using Min_Amount / Max_Amount
        for (const s of stages) {
          if (totalAmount >= s.thresholdMin && (totalAmount <= s.thresholdMax || s.thresholdMax === Infinity)) {
            currentStage = s.stage;
            currentApprover = s.approver;
          }
        }
      }

      // Safety: always ensure we have at least a Stage 1 approver if matrix has data
      // (prevents "current approver not picked" when amount is 0 or doesn't match any threshold)
      if (!currentApprover && stages.length > 0) {
        const first = stages[0];
        currentStage = first.stage;
        currentApprover = first.approver;
      }

      // Build initial approval history with threshold decision for tracker visibility
      const historyEntry: any = {
        stage: currentStage,
        approver: currentApprover || '',
        action: 'Pending',
        date: now,
        comment: ''
      };

      if (currentStage === 2) {
        historyEntry.comment = `Amount exceeded Stage 1 Max_Amount threshold. Escalated to CFO (Stage 2) per Approval Matrix.`;
      } else {
        historyEntry.comment = `Stage 1 approval assigned based on Total PO Amount within threshold.`;
      }

      const initialHistory = JSON.stringify([historyEntry]);

      // Primary payload (matches the documented required columns + what MR/PR inits write)
      const updateData: any = {
        Status: 'Pending Approval',
        Approval_Status: 'Pending',
        Approval_Level: currentStage,
        Approval_Started_On: now,
        Approval_History: initialHistory,
      };
      if (currentApprover) {
        updateData.Current_Approver = currentApprover;
      }

      try {
        await this.updateListItem('PRC_Purchase_Order_Register', itemId, updateData);
      } catch (primaryErr) {
        // Schema on PRC_Purchase_Order_Register may not yet have the full approval columns
        // (Approval_Level / Approval_Started_On / Approval_History / Current_Approver / Status / Approval_Status may be missing or have different internal names).
        // We swallow the error so the PO creation itself is never blocked.
        console.warn('[PO Approval] Could not write approval metadata to PO list (columns may not exist yet). PO was created successfully. Error was:', primaryErr);
        // Do not re-throw – this is best-effort metadata for the tracker.
      }
    } catch (error) {
      console.error('Error initializing PO approval:', error);
    }
  }

  private extractApproverName(approverField: any): string | null {
    if (!approverField) return null;
    if (typeof approverField === 'object') {
      return approverField.Title || approverField.Email || approverField.LoginName || null;
    }
    return String(approverField);
  }

  /**
   * Mark Purchase Order as fully approved
   */
  public async finalizePurchaseOrderApproval(itemId: number): Promise<void> {
    try {
      const now = new Date().toISOString();
      await this.updateListItem('PRC_Purchase_Order_Register', itemId, {
        Approval_Status: 'Approved',
        Approval_Completed_On: now,
        Status: 'Approved',
      });
    } catch (error) {
      console.error('Error finalizing PO approval:', error);
    }
  }

  /**
   * Trigger PDF generation for Purchase Order (via Power Automate or Azure Function)
   * Now supports multi-PR: pass linkedPrIds so the generated PDF can include line items from all selected PRs.
   */
  public async triggerPurchaseOrderPdfGeneration(
    poNumber: string,
    poId: number,
    linkedPrIds: number[] = []
  ): Promise<void> {
    try {
      // This would call a Power Automate HTTP trigger or Azure Function
      // Example endpoint (to be configured):
      // const endpoint = 'https://prod-xxx.azurewebsites.net/api/GeneratePO';

      const payload = {
        poNumber,
        poId,
        linkedPrIds,                    // Array of PR IDs for consolidated PDF with multiple PR line items
        generatedAt: new Date().toISOString(),
      };

      console.log(`[PDF] Triggering PDF generation for PO: ${poNumber} (ID: ${poId}) with ${linkedPrIds.length} linked PR(s)`, payload);

      // Placeholder for actual HTTP call (uncomment and configure when ready):
      // await this.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      //   headers: { 'Content-Type': 'application/json' },
      //   body: JSON.stringify(payload)
      // });
    } catch (error) {
      console.error('Error triggering PO PDF generation:', error);
    }
  }

  /**
   * Initialize GRN approval workflow
   */
  /**
   * Initialize GRN approval workflow (1-stage)
   * Looks up SYS_Approval_Matrix where Module = 'Good Receipt Note' and Stage = 1
   */
  public async initializeGoodsReceivedNoteApproval(itemId: number): Promise<void> {
    try {
      const now = new Date().toISOString();

      let currentApprover: string | null = null;
      try {
        const approvalMatrix = await this.getListData(
          'SYS_Approval_Matrix',
          `Module eq 'Good Receipt Note' and Stage eq 1`,
          1,
          0,
          'Approver_User'
        );
        if (approvalMatrix.length > 0) {
          const approverField = approvalMatrix[0].Approver_User || approvalMatrix[0].Approver;
          if (typeof approverField === 'object' && approverField !== null) {
            currentApprover = approverField.Title || approverField.Email || null;
          } else {
            currentApprover = approverField ? String(approverField) : null;
          }
        }
      } catch (e) {
        console.warn('Approval Matrix lookup failed for GRN (Good Receipt Note)');
      }

      const updateData: any = {
        Status: 'Pending Approval',
        Approval_Status: 'Pending',
        Approval_Level: 1,
        Approval_Started_On: now,
      };

      if (currentApprover) {
        updateData.Current_Approver = currentApprover;
      }

      // Write initial history so tracker shows Stage 1
      const history = JSON.stringify([{
        stage: 1,
        approver: currentApprover || '',
        action: 'Pending',
        date: now,
        comment: 'GRN approval assigned from Approval Matrix (Good Receipt Note, Stage 1)'
      }]);
      updateData.Approval_History = history;

      try {
        await this.updateListItem('PRC_GRN_Register', itemId, updateData);
      } catch (updateErr) {
        // GRN list may not have the full approval columns yet.
        // Make it non-fatal so GRN creation succeeds (same as PO behavior).
        console.warn('[GRN Approval] Could not write approval metadata (columns may be missing). GRN created successfully. Error:', updateErr);
      }
    } catch (error) {
      console.error('Error initializing GRN approval:', error);
    }
  }

  /**
   * Process approve/reject for Goods Received Note (1-stage GRN approval).
   * On successful approval of Stage 1, it records "Notify Finance for Payment"
   * and prepares the data for the UI to send email to the Finance Manager
   * (looked up from Payment Request matrix).
   */
  public async processGoodsReceivedNoteApprovalAction(
    itemId: number,
    action: 'approve' | 'reject',
    approverEmail: string,
    comment: string = ''
  ): Promise<{ financeManagerEmail?: string | null; poId?: number }> {
    const now = new Date().toISOString();
    const items = await this.getListData('PRC_GRN_Register', `ID eq ${itemId}`, 1);
    if (!items.length) return {};

    const item = items[0];
    let history: any[] = [];
    try { history = JSON.parse(item.Approval_History || '[]'); } catch {}

    const poId = item.PO_NumberId || item.PO_Number || item.Related_PO_ID;

    if (action === 'reject') {
      history.push({ stage: 1, approver: approverEmail, action: 'Rejected', date: now, comment });
      await this.updateListItem('PRC_GRN_Register', itemId, {
        Approval_Status: 'Rejected',
        Status: 'Rejected',
        Approval_History: JSON.stringify(history)
      });
      return { poId };
    }

    // Approve path - GRN is 1-stage only
    const currentLevel = 1;
    const lastEntry = history[history.length - 1];
    if (lastEntry && lastEntry.action === 'Pending') {
      lastEntry.action = 'Approved';
      lastEntry.approver = approverEmail;
      lastEntry.date = now;
      lastEntry.comment = comment;
    } else {
      history.push({ stage: 1, approver: approverEmail, action: 'Approved', date: now, comment });
    }

    // Finalize GRN as Approved
    await this.updateListItem('PRC_GRN_Register', itemId, {
      Approval_Status: 'Approved',
      Approval_Completed_On: now,
      Status: 'Approved',
      Current_Approver: null,
      Approval_History: JSON.stringify(history)
    });

    // Trigger Finance notification (records "Notify Finance for Payment" + updates PO)
    let financeManagerEmail: string | null = null;
    if (poId) {
      try {
        financeManagerEmail = await this.notifyFinanceManagerAfterGRN(itemId, poId);
      } catch (e) {
        console.error('Error during GRN Finance notification step:', e);
      }
    }

    return { financeManagerEmail, poId };
  }

  /**
   * Update Purchase Order after GRN creation
   */
  public async updatePOAfterGRN(poId: number): Promise<void> {
    try {
      await this.updateListItem('PRC_Purchase_Order_Register', poId, {
        Delivery_Status: 'Received',
        Status: 'Received',
      });
    } catch (error) {
      console.error('Error updating PO after GRN:', error);
    }
  }

  /**
   * Notify Finance Manager (looked up from Payment Request matrix where Approver_Role contains "Finance Manager")
   * This is used as the "second stage" for GRN flow (per user requirement).
   * The tracker for GRN will display "Notify Finance for Payment".
   */
  public async notifyFinanceManagerAfterGRN(grnId: number, poId: number): Promise<string | null> {
    let financeManager: string | null = null;
    const matrix = await this.getListData(
      'SYS_Approval_Matrix',
      `Module eq 'Payment Request'`,
      10,
      0,
      'Approver_User'
    );

    for (const row of matrix) {
      const role = (row.Approver_Role || '').toString().toLowerCase();
      if (role.includes('finance manager') || role.includes('financemanager')) {
        const f = row.Approver_User || row.Approver;
        if (typeof f === 'object' && f) {
          financeManager = f.Email || f.Title || f.LoginName || null;
        } else if (f) {
          financeManager = String(f);
        }
        break;
      }
    }

    if (!financeManager && matrix.length > 0) {
      const f = matrix[0].Approver_User || matrix[0].Approver;
      if (typeof f === 'object' && f) {
        financeManager = f.Email || f.Title || null;
      } else {
        financeManager = f ? String(f) : null;
      }
    }

    try {
      const now = new Date().toISOString();
      const historyEntry = {
        stage: 2,
        approver: financeManager || 'Finance',
        action: 'Notify Finance',
        date: now,
        comment: 'GRN approved. Finance Manager notified for Payment Request eligibility.'
      };

      try {
        await this.updateListItem('PRC_GRN_Register', grnId, {
          Approval_Status: 'Approved',
          Status: 'Approved',
          Current_Approver: null,
          Approval_Completed_On: now,
          Approval_History: JSON.stringify([historyEntry]),
        });
      } catch {}

      try {
        await this.updateListItem('PRC_Purchase_Order_Register', poId, {
          Delivery_Status: 'Received',
          Payment_Eligible: true,
          GRN_Completed_On: now,
        });
      } catch {}

      console.log(`[GRN Finance Notification] Finance Manager: ${financeManager || 'Not found'}`);
      return financeManager;
    } catch (error) {
      console.warn('[GRN] Finance notification step skipped (matrix columns may be missing).');
      return null;
    }
  }
  /**
   * Legacy / simple notification (kept for backward compatibility)
   */
  public async notifyFinanceForPayment(poId: number, grnId: number): Promise<void> {
    try {
      await this.updateListItem('PRC_Purchase_Order_Register', poId, {
        Payment_Eligible: true,
        GRN_Completed_On: new Date().toISOString(),
        Delivery_Status: 'Received',
      });
      console.log(`[Finance] PO ${poId} is now eligible for Payment Request (GRN: ${grnId})`);
    } catch (error) {
      console.error('Error notifying Finance:', error);
    }
  }

  /**
   * Initialize Payment Request approval with multi-stage flow
   * (Manager → Finance Lead → Director based on amount thresholds)
   */
  public async initializePaymentRequestApproval(itemId: number, amount: number = 0): Promise<void> {
    try {
      const now = new Date().toISOString();

      const approvalMatrix = await this.getListData(
        'SYS_Approval_Matrix',
        `Module eq 'Payment Request'`,
        10,
        0,
        'Approver_User'
      );

      const stages = approvalMatrix
        .map(item => ({
          stage: item.Stage || 1,
          approver: this.extractApproverName(item.Approver_User || item.Approver),
          threshold: item.Threshold_Max || 0,
        }))
        .sort((a, b) => a.stage - b.stage);

      let currentStage = 1;
      let currentApprover: string | null = null;

      for (const stage of stages) {
        if (amount >= stage.threshold) {
          currentStage = stage.stage;
          currentApprover = stage.approver;
        }
      }

      if (!currentApprover && stages.length > 0) {
        currentApprover = stages[0].approver;
        currentStage = stages[0].stage;
      }

      const updateData: any = {
        Payment_Status: 'Pending',
        Approval_Level: currentStage,
        Approval_Started_On: now,
      };

      if (currentApprover) {
        updateData.Current_Approver = currentApprover;
      }

      await this.updateListItem('FIN_Payment_Request_Register', itemId, updateData);
    } catch (error) {
      console.error('Error initializing Payment Request approval:', error);
    }
  }

  /**
   * Finalize Payment Request approval
   */
  public async finalizePaymentRequestApproval(itemId: number): Promise<void> {
    try {
      const now = new Date().toISOString();
      await this.updateListItem('FIN_Payment_Request_Register', itemId, {
        Payment_Status: 'Approved',
        Approval_Completed_On: now,
        Status: 'Approved',
      });
    } catch (error) {
      console.error('Error finalizing Payment Request:', error);
    }
  }

  /**
   * Validate that a GRN exists for the linked PO
   */
  public async validateGRNExists(poId: number): Promise<boolean> {
    try {
      const grns = await this.getListData(
        'PRC_Goods_Received_Note',
        `PO_NumberId eq ${poId}`,
        1
      );
      return grns.length > 0;
    } catch (error) {
      console.error('Error validating GRN existence:', error);
      return false;
    }
  }
}
