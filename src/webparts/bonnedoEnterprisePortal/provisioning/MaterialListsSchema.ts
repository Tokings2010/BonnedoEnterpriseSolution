/**
 * SharePoint list schema reference for the Material Management module.
 * Create these lists manually in SharePoint or via PnP PowerShell before using the module.
 */
export interface ISharePointColumnDefinition {
  internalName: string;
  displayName: string;
  type: string;
  required: boolean;
  choices?: string[];
  notes?: string;
}

export interface ISharePointListSchema {
  listName: string;
  description: string;
  columns: ISharePointColumnDefinition[];
}

export const MATERIAL_LIST_SCHEMAS: ISharePointListSchema[] = [
  {
    listName: 'ENT_Materials_Master',
    description: 'Material catalog master data',
    columns: [
      { internalName: 'Title', displayName: 'Title', type: 'Text', required: true },
      { internalName: 'Material_Code', displayName: 'Material Code', type: 'Text', required: true, notes: 'Unique; format MAT-####' },
      { internalName: 'Material_Name', displayName: 'Material Name', type: 'Text', required: true },
      { internalName: 'Category', displayName: 'Category', type: 'Choice', required: true, choices: ['PIP', 'VAL', 'FIT', 'RED', 'FLG', 'TEE', 'BLT', 'NUT', 'OEM', 'CON', 'PPE', 'Others'] },
      { internalName: 'SubType', displayName: 'Sub Type', type: 'Choice', required: false, choices: ['Carbon Steel Line', 'Stainless Steel', 'GRE', 'API 600', 'Class 150', 'Long Radius BW', 'Concentric BW', '150# RF', 'Equal BW', 'Grade 8.8', 'OEM Module', 'Consumable', 'PPE', 'N/A'] },
      { internalName: 'Size', displayName: 'Size', type: 'Choice', required: false, choices: ['2"', '3"', '4"', '6"', '8"', '10"', '12"', '14"', '16"', 'M16', '3.2mm', 'N/A'] },
      { internalName: 'UOM', displayName: 'Unit of Measure', type: 'Choice', required: true, choices: ['EA', 'JT', 'M', 'KG', 'SET', 'LOT', 'BOX'] },
      { internalName: 'Standard_Cost', displayName: 'Standard Cost', type: 'Currency', required: false },
      { internalName: 'MinStockLevel', displayName: 'Min Stock Level', type: 'Number', required: false },
      { internalName: 'QRCodeURL', displayName: 'QR Code URL', type: 'Hyperlink', required: false },
      { internalName: 'Specification', displayName: 'Specification', type: 'Text', required: false },
      { internalName: 'Active', displayName: 'Active', type: 'Boolean', required: false, notes: 'Default: Yes' },
    ],
  },
  {
    listName: 'ENT_Warehouses_Master',
    description: 'Warehouse master data',
    columns: [
      { internalName: 'Title', displayName: 'Title', type: 'Text', required: true },
      { internalName: 'WarehouseCode', displayName: 'Warehouse Code', type: 'Text', required: true, notes: 'Unique; joined to inventory Location' },
      { internalName: 'Warehouse_Name', displayName: 'Warehouse Name', type: 'Text', required: true },
      { internalName: 'Location', displayName: 'Physical Location', type: 'Text', required: false },
      { internalName: 'Status', displayName: 'Status', type: 'Choice', required: true, choices: ['Active', 'Inactive'] },
    ],
  },
  {
    listName: 'ENT_Inventory_Register',
    description: 'Stock on hand by warehouse location',
    columns: [
      { internalName: 'Title', displayName: 'Title', type: 'Text', required: true },
      { internalName: 'Material_Code', displayName: 'Material Code', type: 'Text', required: true, notes: 'Do not use field_1' },
      { internalName: 'Location', displayName: 'Warehouse', type: 'Text', required: true, notes: 'Stores WarehouseCode' },
      { internalName: 'Project_Code', displayName: 'Project Code', type: 'Text', required: false },
      { internalName: 'Qty_On_Hand', displayName: 'Qty On Hand', type: 'Number', required: true },
      { internalName: 'QtyReserved', displayName: 'Qty Reserved', type: 'Number', required: false },
      { internalName: 'BinLocation', displayName: 'Bin Location', type: 'Text', required: false },
      { internalName: 'BatchNumber', displayName: 'Batch Number', type: 'Text', required: false },
      { internalName: 'Condition', displayName: 'Condition', type: 'Choice', required: false, choices: ['Good', 'Damaged', 'Quarantine'] },
      { internalName: 'Status', displayName: 'Status', type: 'Choice', required: false, choices: ['Available', 'Reserved', 'Hold'] },
      { internalName: 'Last_Movement_Date', displayName: 'Last Movement Date', type: 'DateTime', required: false },
      { internalName: 'DateReceived', displayName: 'Date Received', type: 'DateTime', required: false },
    ],
  },
  {
    listName: 'ENT_Inventory_Movements_Register',
    description: 'Inventory movement audit trail',
    columns: [
      { internalName: 'Title', displayName: 'Title', type: 'Text', required: true },
      { internalName: 'Movement_Type', displayName: 'Movement Type', type: 'Choice', required: true, choices: ['GRN', 'Transfer Out', 'Issue', 'Return', 'Scrap'] },
      { internalName: 'Material_Code', displayName: 'Material Code', type: 'Text', required: true, notes: 'Do not use field_2' },
      { internalName: 'Qty', displayName: 'Quantity', type: 'Number', required: true },
      { internalName: 'From_x0020_Location', displayName: 'From Location', type: 'Text', required: false },
      { internalName: 'To_x0020_Location', displayName: 'To Location', type: 'Text', required: false },
      { internalName: 'Project_Code', displayName: 'Project Code', type: 'Text', required: false },
      { internalName: 'Note', displayName: 'Note', type: 'Note', required: false },
    ],
  },
];
