import { IListItem } from '../services/SharePointService';
import {
  IMaterialMasterRecord,
  IInventoryRecord,
  IInventoryMovementRecord,
  IWarehouseRecord,
} from '../models/DataModels';

interface IHyperlinkFieldValue {
  Url?: string;
  Description?: string;
}

function getHyperlinkUrl(value: unknown): string {
  if (!value) {
    return '';
  }

  if (typeof value === 'string') {
    return value;
  }

  if (typeof value === 'object') {
    const hyperlink = value as IHyperlinkFieldValue;
    return hyperlink.Url || hyperlink.Description || '';
  }

  return '';
}

/** Resolves material code from ENT_ column or legacy field_1 / field_2. */
export function resolveMaterialCode(item: IListItem, legacyField?: string): string {
  if (item.Material_Code) {
    return String(item.Material_Code);
  }
  if (legacyField && item[legacyField]) {
    return String(item[legacyField]);
  }
  return '';
}

export function mapMaterialMasterRecord(item: IListItem): IMaterialMasterRecord {
  return {
    ID: item.ID,
    Title: item.Title || '',
    Material_Code: item.Material_Code || '',
    Material_Name: item.Material_Name || item.Title || '',
    Category: item.Category || '',
    SubType: item.SubType || '',
    Size: item.Size || '',
    UOM: item.UOM || '',
    Standard_Cost: item.Standard_Cost || 0,
    MinStockLevel: item.MinStockLevel || 0,
    QRCodeURL:
      getHyperlinkUrl(item.QRCodeURL) ||
      getHyperlinkUrl(item.qrcodeurl) ||
      getHyperlinkUrl(item.MaterialQRCode) ||
      getHyperlinkUrl(item.Material_QR_Code) ||
      '',
    Specification: item.Specification || '',
    Active: item.Active !== false,
  };
}

export function mapWarehouseRecord(item: IListItem): IWarehouseRecord {
  return {
    ID: item.ID,
    Title: item.Title || '',
    WarehouseCode: item.WarehouseCode || item.Title || '',
    Warehouse_Name: item.Warehouse_Name || '',
    Location: item.Location || '',
    Status: item.Status || 'Active',
  };
}

export function mapInventoryRecord(item: IListItem): IInventoryRecord {
  return {
    ID: item.ID,
    Title: item.Title || '',
    Material_Code: resolveMaterialCode(item, 'field_1'),
    Location: item.Location || '',
    Project_Code: item.Project_Code || '',
    Qty_On_Hand: item.Qty_On_Hand || 0,
    QtyReserved: item.QtyReserved || 0,
    Last_Movement_Date: item.Last_Movement_Date || '',
    BinLocation: item.BinLocation || '',
    BatchNumber: item.BatchNumber || '',
    Condition: item.Condition || 'Good',
    DateReceived: item.DateReceived || '',
    Status: item.Status || 'Available',
  };
}

export function mapInventoryMovementRecord(item: IListItem): IInventoryMovementRecord {
  return {
    ID: item.ID,
    Title: item.Title || '',
    Movement_Type: item.Movement_Type || '',
    Material_Code: resolveMaterialCode(item, 'field_2'),
    Qty: item.Qty || 0,
    From_Location: item.From_Location || item.From_x0020_Location || '',
    To_Location: item.To_Location || item.To_x0020_Location || '',
    Project_Code: item.Project_Code || '',
    Note: item.Note || '',
    Created: item.Created || '',
  };
}
