/**
 * Common data models for the Enterprise Portal
 */

export interface IListItem {
  ID: number;
  Title?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

export interface IDataGridColumn {
  key: string;
  name: string;
  fieldName: string;
  minWidth?: number;
  maxWidth?: number;
  isResizable?: boolean;
  isSorted?: boolean;
  isSortedDescending?: boolean;
  onColumnClick?: (column: IDataGridColumn) => void;
}

export interface IFinanceRecord extends IListItem {
  Title: string;
  Amount?: number;
  Status?: string;
  DueDate?: string;
  Approver?: string;
  Description?: string;
}

export interface IProcurementRecord extends IListItem {
  Title: string;
  Vendor?: string;
  Amount?: number;
  Status?: string;
  RequestDate?: string;
  DeliveryDate?: string;
}

export interface IProjectRecord extends IListItem {
  Title: string;
  ProjectManager?: string;
  Status?: string;
  StartDate?: string;
  EndDate?: string;
  Budget?: number;
  Progress?: number;
}

export interface IDashboardMetrics {
  totalApprovals: number;
  pendingApprovals: number;
  totalInvoices: number;
  totalAmount: number;
  averageProcessingTime: number;
}

export interface IApprovalRecord extends IListItem {
  Title: string;
  InvoiceNumber?: string;
  Vendor?: string;
  Amount?: number;
  FundType?: string;
  Approval_Status?: string;
  Current_Approver?: string;
  Approval_History?: string;
  Comments?: string;
}

export interface IApprovalAction {
  action: 'approve' | 'reject';
  comment: string;
  timestamp: string;
  approver: string;
}

// Procurement Models
export interface IMaterialRequest extends IListItem {
  Title: string;
  Project_Code?: string;
  Material?: string;
  Quantity?: number;
  UOM?: string;
  Request_Date?: string;
  Status?: string;
  Approval_Status?: string;
  QR_Code?: string;
  Notes?: string;
}

export interface IPurchaseRequisition extends IListItem {
  Title: string;
  Project_Code?: string;
  Description?: string;
  Quantity?: number;
  UOM?: string;
  EstimatedCost?: number;
  Request_Date?: string;
  Status?: string;
  Approval_Status?: string;
  QR_Code?: string;
}

export interface IPurchaseOrder extends IListItem {
  Title: string;
  Vendor?: string;
  Description?: string;
  Quantity?: number;
  UOM?: string;
  UnitPrice?: number;
  TotalAmount?: number;
  DeliveryDate?: string;
  Status?: string;
  Approval_Status?: string;
  QR_Code?: string;
}

export interface IGoodsReceivedNote extends IListItem {
  Title: string;
  PO_Number?: string;
  Vendor?: string;
  Quantity_Received?: number;
  UOM?: string;
  Received_Date?: string;
  Status?: string;
  QR_Code?: string;
  Notes?: string;
}

export interface IProcurementFormData {
  Project_Code: string;
  Material?: string;
  Description?: string;
  Quantity: number;
  UOM: string;
  EstimatedCost?: number;
  Vendor?: string;
  UnitPrice?: number;
  Notes?: string;
}
