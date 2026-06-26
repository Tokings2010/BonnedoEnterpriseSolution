import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';
import { IMaterialMasterRecord } from '../models/DataModels';
import { logCostTransaction } from './CostLinkService';

// ============================================
// Inventory Movement Processing Interfaces
// ============================================

export interface IInventoryMovementRequest {
  /** Movement type exactly as stored in SharePoint */
  movementType: 'GRN' | 'Transfer Out' | 'Issue' | 'Return' | 'Scrap';
  /** Material code being moved */
  materialCode: string;
  /** Quantity moved */
  quantity: number;
  /** Source warehouse (required for Issue, Scrap, Transfer Out) */
  fromLocation?: string;
  /** Destination warehouse (required for GRN, Transfer Out, Return) */
  toLocation?: string;
  /** Optional project code */
  projectCode?: string;
  /** Optional user note */
  note?: string;
  /** ISO timestamp or empty for current time */
  timestamp?: string;
}

export interface IMovementProcessingResult {
  success: boolean;
  /** ID of the created/updated inventory record(s) */
  inventoryIds: number[];
  /** Human-readable summary */
  summary: string;
  /** Error details if failed */
  error?: string;
}

export interface IMovementProgress {
  step: string;
  message: string;
  progress: number;
}

// ============================================
// Inventory Movement Service
// ============================================

/**
 * InventoryMovementService — replaces Power Automate flows for
 * inventory movement processing.
 *
 * Flow trigger: When an item is created on ENT_Inventory_Movements_Register
 *
 * The service processes each movement type inline without polling,
 * providing synchronous results, detailed error messages, and
 * automatic cost transaction logging.
 */
export class InventoryMovementService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;

  constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
    this.spHttpClient = spHttpClient;
    this.pageContext = pageContext;
  }

  // ============================================
  // PUBLIC: Process a movement (main entry point)
  // ============================================

  /**
   * Process an inventory movement: validates the request, applies the
   * correct business logic per movement type, updates the Inventory
   * register, and logs cost transactions for Issues.
   *
   * This single method replaces an entire Power Automate flow with
   * switch-case branching.
   */
  public async processMovement(
    request: IInventoryMovementRequest,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    const timestamp = request.timestamp || new Date().toISOString();
    const code = request.materialCode.trim();
    const qty = request.quantity;
    const fromLoc = (request.fromLocation || '').trim();
    const toLoc = (request.toLocation || '').trim();

    // ── Pre-validation ──
    if (!code) {
      return { success: false, inventoryIds: [], summary: 'Material code is required.', error: 'Material_Code cannot be empty.' };
    }
    if (qty <= 0 || !Number.isFinite(qty)) {
      return { success: false, inventoryIds: [], summary: 'Quantity must be positive.', error: `Invalid quantity: ${qty}` };
    }

    onProgress?.({ step: 'validate', message: `Validating movement: ${request.movementType} — ${code}`, progress: 10 });

    // ── Verify material exists in master ──
    const material = await this.fetchMaterialMaster(code);
    if (!material) {
      return { success: false, inventoryIds: [], summary: `Material ${code} not found in master.`, error: `Material_Code '${code}' does not exist in ${SHAREPOINT_LISTS.MATERIALS_MASTER}.` };
    }
    if (material.Active === false) {
      return { success: false, inventoryIds: [], summary: `Material ${code} is inactive.`, error: `Material '${code}' is inactive. Reactivate before recording movements.` };
    }

    onProgress?.({ step: 'processing', message: `Processing ${request.movementType} for ${code} (qty: ${qty})`, progress: 30 });

    // ── Route to the correct handler ──
    try {
      switch (request.movementType) {
        case 'GRN':
          return await this.handleGrn(code, qty, toLoc, request.projectCode, timestamp, onProgress);
        case 'Transfer Out':
          return await this.handleTransferOut(code, qty, fromLoc, toLoc, request.projectCode, timestamp, onProgress);
        case 'Issue':
          return await this.handleIssue(code, qty, fromLoc, request.projectCode, timestamp, material.Standard_Cost || 0, onProgress);
        case 'Return':
          return await this.handleReturn(code, qty, toLoc, request.projectCode, timestamp, onProgress);
        case 'Scrap':
          return await this.handleScrap(code, qty, fromLoc, timestamp, onProgress);
        default:
          return { success: false, inventoryIds: [], summary: `Unknown movement type: ${request.movementType}`, error: `Unsupported Movement_Type: '${request.movementType}'` };
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Unknown error during movement processing';
      return { success: false, inventoryIds: [], summary: `Processing failed: ${errMsg}`, error: errMsg };
    }
  }

  // ============================================
  // CASE: GRN (Goods Received Note)
  //   → Create item in Inventory
  //   → field_1 = Material_Code, Location = To_Location
  //   → Qty_On_Hand = Qty, Status = "Available"
  // ============================================

  private async handleGrn(
    code: string,
    qty: number,
    toLocation: string,
    projectCode: string | undefined,
    timestamp: string,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    if (!toLocation) {
      return { success: false, inventoryIds: [], summary: 'Destination warehouse is required for GRN.', error: 'To_Location cannot be empty for GRN.' };
    }

    onProgress?.({ step: 'grn-create', message: `Creating inventory record at ${toLocation}...`, progress: 60 });

    // Try to find existing inventory record for this code + location + project
    const existing = await this.findInventoryRecord(code, toLocation, projectCode);

    if (existing) {
      // Accumulate quantity
      const currentQty = Number(existing.Qty_On_Hand || 0);
      const newQty = currentQty + qty;
      await this.updateInventoryRecord(existing.ID, {
        Qty_On_Hand: newQty,
        Last_Movement_Date: timestamp,
        Status: 'Available',
      });
      return {
        success: true,
        inventoryIds: [existing.ID],
        summary: `GRN: Added ${qty} to ${code} at ${toLocation} (was ${currentQty}, now ${newQty}).`,
      };
    }

    // Create new inventory record
    const newId = await this.createInventoryRecord({
      Title: `${code} - ${toLocation}`,
      Material_Code: code,
      Location: toLocation,
      Project_Code: projectCode || null,
      Qty_On_Hand: qty,
      QtyReserved: 0,
      Condition: 'Good',
      Status: 'Available',
      DateReceived: timestamp,
      Last_Movement_Date: timestamp,
    });

    return {
      success: true,
      inventoryIds: [newId],
      summary: `GRN: Created inventory record for ${code} at ${toLocation} (qty: ${qty}).`,
    };
  }

  // ============================================
  // CASE: Transfer Out (UPDATED)
  //   Step A: Subtract from source warehouse
  //   Step B: Add to destination warehouse
  //   Sequential (not parallel)
  // ============================================

  private async handleTransferOut(
    code: string,
    qty: number,
    fromLocation: string,
    toLocation: string,
    projectCode: string | undefined,
    timestamp: string,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    if (!fromLocation) {
      return { success: false, inventoryIds: [], summary: 'Source warehouse is required for Transfer Out.', error: 'From_Location cannot be empty for Transfer Out.' };
    }
    if (!toLocation) {
      return { success: false, inventoryIds: [], summary: 'Destination warehouse is required for Transfer Out.', error: 'To_Location cannot be empty for Transfer Out.' };
    }
    if (fromLocation === toLocation) {
      return { success: false, inventoryIds: [], summary: 'Source and destination must be different.', error: 'From_Location and To_Location must differ for Transfer Out.' };
    }

    onProgress?.({ step: 'transfer-stepA', message: `Step A: Subtracting ${qty} from ${fromLocation}...`, progress: 40 });

    // ── Step A: Get items from Inventory (filter by Material_Code + Location = From_Location) ──
    const sourceRecord = await this.findInventoryRecord(code, fromLocation, projectCode);
    if (!sourceRecord) {
      return {
        success: false,
        inventoryIds: [],
        summary: `No inventory found for ${code} at ${fromLocation}.`,
        error: `Get_items_transfer: No record matching Material_Code='${code}' and Location='${fromLocation}'.`,
      };
    }

    const currentQty = Number(sourceRecord.Qty_On_Hand || 0);
    if (currentQty < qty) {
      return {
        success: false,
        inventoryIds: [],
        summary: `Insufficient stock at ${fromLocation}. Available: ${currentQty}, requested: ${qty}.`,
        error: `Insufficient stock for Transfer Out. Available: ${currentQty}, needed: ${qty}.`,
      };
    }

    // Subtract from source
    const newSourceQty = currentQty - qty;
    await this.updateInventoryRecord(sourceRecord.ID, {
      Qty_On_Hand: newSourceQty,
      Last_Movement_Date: timestamp,
    });

    onProgress?.({ step: 'transfer-stepB', message: `Step B: Adding ${qty} to ${toLocation}...`, progress: 70 });

    // ── Step B: Add to destination warehouse ──
    const destRecord = await this.findInventoryRecord(code, toLocation, projectCode);
    let destId: number;

    if (destRecord) {
      const newDestQty = Number(destRecord.Qty_On_Hand || 0) + qty;
      await this.updateInventoryRecord(destRecord.ID, {
        Qty_On_Hand: newDestQty,
        Last_Movement_Date: timestamp,
        Status: 'Available',
      });
      destId = destRecord.ID;
    } else {
      destId = await this.createInventoryRecord({
        Title: `${code} - ${toLocation}`,
        Material_Code: code,
        Location: toLocation,
        Project_Code: projectCode || null,
        Qty_On_Hand: qty,
        QtyReserved: 0,
        Condition: 'Good',
        Status: 'Available',
        DateReceived: timestamp,
        Last_Movement_Date: timestamp,
      });
    }

    return {
      success: true,
      inventoryIds: [sourceRecord.ID, destId],
      summary: `Transfer Out: Moved ${qty} of ${code} from ${fromLocation} to ${toLocation}.`,
    };
  }

  // ============================================
  // CASE: Issue (UPDATED — uses From_Location)
  //   Get items → Update item (subtract qty)
  // ============================================

  private async handleIssue(
    code: string,
    qty: number,
    fromLocation: string,
    projectCode: string | undefined,
    timestamp: string,
    standardCost: number,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    if (!fromLocation) {
      return { success: false, inventoryIds: [], summary: 'Source warehouse is required for Issue.', error: 'From_Location cannot be empty for Issue.' };
    }

    onProgress?.({ step: 'issue-get', message: `Fetching inventory at ${fromLocation}...`, progress: 50 });

    // Get items from Inventory: filter by Material_Code + Location = From_Location
    const sourceRecord = await this.findInventoryRecord(code, fromLocation, projectCode);
    if (!sourceRecord) {
      return {
        success: false,
        inventoryIds: [],
        summary: `No inventory found for ${code} at ${fromLocation}.`,
        error: `Get_items_issue: No record matching Material_Code='${code}' and Location='${fromLocation}'.`,
      };
    }

    const currentQty = Number(sourceRecord.Qty_On_Hand || 0);
    if (currentQty < qty) {
      return {
        success: false,
        inventoryIds: [],
        summary: `Insufficient stock at ${fromLocation}. Available: ${currentQty}, requested: ${qty}.`,
        error: `Insufficient stock for Issue. Available: ${currentQty}, needed: ${qty}.`,
      };
    }

    // Subtract qty
    const newQty = currentQty - qty;
    await this.updateInventoryRecord(sourceRecord.ID, {
      Qty_On_Hand: newQty,
      Last_Movement_Date: timestamp,
    });

    // Log cost transaction for material issue
    if (projectCode && standardCost > 0) {
      try {
        await logCostTransaction(this.spHttpClient, this.pageContext, {
          projectCode: projectCode,
          phase: '',
          transactionType: 'Material Issue',
          amount: standardCost * qty,
          referenceId: code,
          referenceType: 'Material Issue',
          description: `Issued ${qty} x ${code} from ${fromLocation}`,
        });
      } catch (costErr) {
        console.warn('[InventoryMovementService] Failed to log cost for Issue:', costErr);
      }
    }

    return {
      success: true,
      inventoryIds: [sourceRecord.ID],
      summary: `Issue: Issued ${qty} of ${code} from ${fromLocation} (remaining: ${newQty}).`,
    };
  }

  // ============================================
  // CASE: Return (unchanged — same as GRN)
  //   Create item in Inventory
  // ============================================

  private async handleReturn(
    code: string,
    qty: number,
    toLocation: string,
    projectCode: string | undefined,
    timestamp: string,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    if (!toLocation) {
      return { success: false, inventoryIds: [], summary: 'Destination warehouse is required for Return.', error: 'To_Location cannot be empty for Return.' };
    }

    onProgress?.({ step: 'return-create', message: `Processing return to ${toLocation}...`, progress: 60 });

    // Same logic as GRN: accumulate if exists
    const existing = await this.findInventoryRecord(code, toLocation, projectCode);
    if (existing) {
      const currentQty = Number(existing.Qty_On_Hand || 0);
      const newQty = currentQty + qty;
      await this.updateInventoryRecord(existing.ID, {
        Qty_On_Hand: newQty,
        Last_Movement_Date: timestamp,
        Status: 'Available',
      });
      return {
        success: true,
        inventoryIds: [existing.ID],
        summary: `Return: Added ${qty} of ${code} at ${toLocation} (was ${currentQty}, now ${newQty}).`,
      };
    }

    const newId = await this.createInventoryRecord({
      Title: `${code} - ${toLocation}`,
      Material_Code: code,
      Location: toLocation,
      Project_Code: projectCode || null,
      Qty_On_Hand: qty,
      QtyReserved: 0,
      Condition: 'Good',
      Status: 'Available',
      DateReceived: timestamp,
      Last_Movement_Date: timestamp,
    });

    return {
      success: true,
      inventoryIds: [newId],
      summary: `Return: Created inventory record for ${code} at ${toLocation} (qty: ${qty}).`,
    };
  }

  // ============================================
  // CASE: Scrap (UPDATED — uses From_Location)
  //   Get items → Update item (subtract qty)
  // ============================================

  private async handleScrap(
    code: string,
    qty: number,
    fromLocation: string,
    timestamp: string,
    onProgress?: (progress: IMovementProgress) => void
  ): Promise<IMovementProcessingResult> {
    if (!fromLocation) {
      return { success: false, inventoryIds: [], summary: 'Source warehouse is required for Scrap.', error: 'From_Location cannot be empty for Scrap.' };
    }

    onProgress?.({ step: 'scrap-get', message: `Fetching inventory at ${fromLocation}...`, progress: 50 });

    const sourceRecord = await this.findInventoryRecord(code, fromLocation, undefined);
    if (!sourceRecord) {
      return {
        success: false,
        inventoryIds: [],
        summary: `No inventory found for ${code} at ${fromLocation}.`,
        error: `Get_items_scrap: No record matching Material_Code='${code}' and Location='${fromLocation}'.`,
      };
    }

    const currentQty = Number(sourceRecord.Qty_On_Hand || 0);
    if (currentQty < qty) {
      return {
        success: false,
        inventoryIds: [],
        summary: `Insufficient stock at ${fromLocation}. Available: ${currentQty}, requested: ${qty}.`,
        error: `Insufficient stock for Scrap. Available: ${currentQty}, needed: ${qty}.`,
      };
    }

    const newQty = currentQty - qty;
    await this.updateInventoryRecord(sourceRecord.ID, {
      Qty_On_Hand: newQty,
      Last_Movement_Date: timestamp,
      Condition: newQty > 0 ? undefined : 'Scrapped',
      Status: newQty > 0 ? 'Available' : 'Scrapped',
    });

    return {
      success: true,
      inventoryIds: [sourceRecord.ID],
      summary: `Scrap: Removed ${qty} of ${code} from ${fromLocation} (remaining: ${newQty}).`,
    };
  }

  // ============================================
  // PRIVATE: SharePoint data access helpers
  // ============================================

  /**
   * Fetch material master record by code.
   */
  private async fetchMaterialMaster(code: string): Promise<IMaterialMasterRecord | null> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const escaped = code.replace(/'/g, "''");
    const url = `${webUrl}/_api/web/lists/getByTitle('${SHAREPOINT_LISTS.MATERIALS_MASTER}')/items?$filter=Material_Code eq '${escaped}'&$top=1&$select=Material_Code,Material_Name,Active,Standard_Cost`;
    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return null;
      const data = await resp.json();
      return data.value?.[0] || null;
    } catch {
      return null;
    }
  }

  /**
   * Find an inventory record by material code + location (+ optional project).
   */
  private async findInventoryRecord(
    materialCode: string,
    location: string,
    projectCode?: string
  ): Promise<{ ID: number; Qty_On_Hand?: number; Location?: string; Material_Code?: string; Status?: string } | null> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const escapedCode = materialCode.replace(/'/g, "''");
    const escapedLoc = location.replace(/'/g, "''");
    let filter = `Material_Code eq '${escapedCode}' and Location eq '${escapedLoc}'`;
    if (projectCode) {
      const escapedProj = projectCode.replace(/'/g, "''");
      filter += ` and Project_Code eq '${escapedProj}'`;
    }
    const url = `${webUrl}/_api/web/lists/getByTitle('${SHAREPOINT_LISTS.INVENTORY_REGISTER}')/items?$filter=${encodeURIComponent(filter)}&$top=1&$select=ID,Qty_On_Hand,Location,Material_Code,Status`;
    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return null;
      const data = await resp.json();
      return data.value?.[0] || null;
    } catch {
      return null;
    }
  }

  /**
   * Create a new inventory record.
   */
  private async createInventoryRecord(data: Record<string, unknown>): Promise<number> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${SHAREPOINT_LISTS.INVENTORY_REGISTER}')/items`;
    const resp = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      },
      body: JSON.stringify(data),
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Failed to create inventory record: ${resp.status} — ${errText}`);
    }
    const result = await resp.json();
    return result.ID;
  }

  /**
   * Update an existing inventory record (partial MERGE).
   */
  private async updateInventoryRecord(id: number, data: Record<string, unknown>): Promise<void> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${SHAREPOINT_LISTS.INVENTORY_REGISTER}')/items(${id})`;
    const resp = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(data),
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Failed to update inventory record ${id}: ${resp.status} — ${errText}`);
    }
  }
}

// ============================================
// Factory function
// ============================================

export function createInventoryMovementService(
  spHttpClient: SPHttpClient,
  pageContext: PageContext
): InventoryMovementService {
  return new InventoryMovementService(spHttpClient, pageContext);
}
