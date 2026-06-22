/**
 * Centralized SharePoint list names for the Material Management module.
 * Use these constants instead of hard-coded strings in components and services.
 */
export const SHAREPOINT_LISTS = {
  MATERIALS_MASTER: 'ENT_Materials_Master',
  WAREHOUSES_MASTER: 'ENT_Warehouses_Master',
  INVENTORY_REGISTER: 'ENT_Inventory_Register',
  INVENTORY_MOVEMENTS_REGISTER: 'ENT_Inventory_Movements_Register',
} as const;

export type SharePointListName = typeof SHAREPOINT_LISTS[keyof typeof SHAREPOINT_LISTS];
