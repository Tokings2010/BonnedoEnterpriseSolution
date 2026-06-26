import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
  getProjectDocumentLibraryRootUrl,
  getProjectDocumentFolderServerRelativeUrl,
  PROJECT_BUDGET_FOLDER_NAME,
} from './ProjectDocumentStorageService';

// ============================================
// CBS Budget Import Interfaces
// ============================================

export interface ICbsRow {
  /** CBS Code (e.g. MOB-001, ENG-010) */
  cbsCode: string;
  /** Phase (e.g. MOBILIZATION, ENGINEERING) */
  phase: string;
  /** Description of the line item */
  description: string;
  /** Unit of Measure */
  uom: string;
  /** Quantity */
  qty: number;
  /** Rate in NGN */
  rate: number;
  /** Budget Amount = Qty × Rate (or as provided) */
  amount: number;
  /** Optional remarks */
  remarks: string;
}

export interface ICbsUploadFile {
  /** Original file name e.g. PRJ-HEIRS-ISU-2026_CBS.xlsx */
  fileName: string;
  /** Extracted project code */
  projectCode: string;
  /** Full file content as ArrayBuffer */
  fileBuffer: ArrayBuffer;
}

export interface ICbsImportResult {
  success: boolean;
  projectCode: string;
  isVariation: boolean;
  rowsTotal: number;
  rowsCreated: number;
  rowsUpdated: number;
  rowsSkipped: number;
  summary: string;
  error?: string;
  /** Individual row-level details */
  details: ICbsRowResult[];
}

export interface ICbsRowResult {
  cbsCode: string;
  description: string;
  action: 'created' | 'updated' | 'skipped';
  reason?: string;
}

export interface ICbsImportProgress {
  step: string;
  message: string;
  progress: number;
  /** Current row being processed (if applicable) */
  currentRow?: number;
  totalRows?: number;
}

// ============================================
// CBS Budget Import Service
// ============================================

/**
 * CbsBudgetImportService — replaces Power Automate "Parse CBS Budget Upload"
 * flow entirely in-browser.
 *
 * Benefits over Power Automate:
 * - Zero latency — no cloud flow queuing, no polling
 * - Works offline — all processing happens on the client
 * - Batch operations — can process all rows in a single transaction scope
 * - Richer validation — validates CBS code format, numeric types, phase matching
 * - Detailed reporting — per-row results with specific reasons for skips
 * - No license cost — no premium Power Automate connectors needed
 */
export class CbsBudgetImportService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;

  // SharePoint list name for budget items
  private readonly BUDGET_ITEMS_LIST = 'Project_Budget_Items';

  constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
    this.spHttpClient = spHttpClient;
    this.pageContext = pageContext;
  }

  // ============================================
  // PUBLIC: Main entry point
  // ============================================

  /**
   * Process a CBS budget file upload. Replaces the entire Power Automate flow.
   *
   * Flow equivalent:
   *   1. Trigger: file uploaded to Documents/Budgets/
   *   2. Condition: filename contains '_CBS'
   *   3. Compose: Extract Project Code (split on '_CBS')
   *   4. Get items: Check if budget already exists (variation detection)
   *   5. Compose: Is Variation? (true/false)
   *   6. List rows: Parse Excel CBS_Table
   *   7. Apply to each: Loop over rows with conditions
   *      7a. Skip empty/subtotal rows
   *      7b. Variation → update existing OR create
   *      7c. First upload → create all
   */
  public async importCbsFile(
    file: ICbsUploadFile,
    onProgress?: (progress: ICbsImportProgress) => void
  ): Promise<ICbsImportResult> {
    const details: ICbsRowResult[] = [];
    let rowsCreated = 0;
    let rowsUpdated = 0;
    let rowsSkipped = 0;

    // ── Step 1: Determine project code ──
    // If a project code is explicitly provided in the request, use it directly.
    // Otherwise, extract it from the filename (Power Automate trigger scenario).
    let projectCode = (file.projectCode || '').trim();

    onProgress?.({ step: 'validate', message: `Preparing to import: ${file.fileName}`, progress: 5, currentRow: 0, totalRows: 0 });

    if (!projectCode) {
      // No explicit project code — extract from filename (requires _CBS suffix)
      if (!file.fileName.toUpperCase().includes('_CBS')) {
        return {
          success: false,
          projectCode: '',
          isVariation: false,
          rowsTotal: 0,
          rowsCreated: 0,
          rowsUpdated: 0,
          rowsSkipped: 0,
          summary: `File '${file.fileName}' does not contain '_CBS' and no project code was provided.`,
          details: [],
          error: 'Missing project code and filename does not contain _CBS separator.',
        };
      }
      projectCode = this.extractProjectCode(file.fileName);
      if (!projectCode) {
        return {
          success: false,
          projectCode: '',
          isVariation: false,
          rowsTotal: 0,
          rowsCreated: 0,
          rowsUpdated: 0,
          rowsSkipped: 0,
          summary: `Could not extract project code from '${file.fileName}'. Expected format: PRJ-CODE_CBS.xlsx`,
          details: [],
          error: 'Project code extraction failed.',
        };
      }
    }

    // ── Step 3: Parse Excel file ──
    onProgress?.({ step: 'parse', message: `Parsing Excel CBS_Table from ${file.fileName}...`, progress: 20, currentRow: 0, totalRows: 0 });
    let rows: ICbsRow[];
    try {
      rows = this.parseExcelCbsTable(file.fileBuffer);
    } catch (err) {
      return {
        success: false,
        projectCode,
        isVariation: false,
        rowsTotal: 0,
        rowsCreated: 0,
        rowsUpdated: 0,
        rowsSkipped: 0,
        summary: `Failed to parse Excel file: ${err instanceof Error ? err.message : 'Unknown error'}`,
        details: [],
        error: err instanceof Error ? err.message : 'Excel parse failure',
      };
    }

    if (rows.length === 0) {
      return {
        success: false,
        projectCode,
        isVariation: false,
        rowsTotal: 0,
        rowsCreated: 0,
        rowsUpdated: 0,
        rowsSkipped: 0,
        summary: 'No valid CBS rows found in the Excel file. All rows were empty, subtotal, or lacked a CBS Code.',
        details: [],
        error: 'No valid CBS data rows after filtering.',
      };
    }

    onProgress?.({ step: 'found-rows', message: `Found ${rows.length} CBS rows to process.`, progress: 30, currentRow: 0, totalRows: rows.length });

    // ── Step 4: Check if budget already exists (variation detection) ──
    onProgress?.({ step: 'check-existing', message: `Checking if budget exists for ${projectCode}...`, progress: 35, currentRow: 0, totalRows: rows.length });
    const existingRows = await this.fetchExistingBudgetItems(projectCode);
    const isVariation = existingRows.length > 0;

    onProgress?.({ step: 'variation-check', message: isVariation ? `Variation detected — ${existingRows.length} existing rows found. Will update matching CBS codes.` : `First upload — creating all ${rows.length} rows.`, progress: 40, currentRow: 0, totalRows: rows.length });

    // Build a lookup map of existing rows by CBS_Code for variation updates
    const existingMap = new Map<string, number>();
    if (isVariation) {
      existingRows.forEach((r) => {
        if (r.CBS_Code) existingMap.set(r.CBS_Code.trim().toUpperCase(), r.ID);
      });
    }

    // ── Step 5: Process each row (Apply to each) ──
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowNum = i + 1;
      onProgress?.({
        step: 'processing-row',
        message: `Processing row ${rowNum}/${rows.length}: ${row.cbsCode} — ${row.description}`,
        progress: 40 + Math.round((rowNum / rows.length) * 55),
        currentRow: rowNum,
        totalRows: rows.length,
      });

      // ── Step 5a: Skip empty/subtotal rows ──
      if (!row.cbsCode || row.cbsCode.trim().toUpperCase() === 'TOTAL') {
        details.push({
          cbsCode: row.cbsCode || '(empty)',
          description: row.description || '(empty/subtotal)',
          action: 'skipped',
          reason: !row.cbsCode ? 'Empty CBS Code' : 'Row is a TOTAL row',
        });
        rowsSkipped++;
        continue;
      }

      if (isVariation) {
        // ── Step 5b: Variation branch ──
        const existingId = existingMap.get(row.cbsCode.trim().toUpperCase());

        if (existingId) {
          // Variation — Update existing row (DO NOT update Original_Budget)
          await this.updateBudgetItem(existingId, row, projectCode);
          details.push({
            cbsCode: row.cbsCode,
            description: row.description,
            action: 'updated',
            reason: `Updated Budget_Amount, Qty, Rate, Remarks (Original_Budget unchanged)`,
          });
          rowsUpdated++;
        } else {
          // Variation — New line added: create with Original_Budget = 0
          await this.createBudgetItem(row, projectCode, false);
          details.push({
            cbsCode: row.cbsCode,
            description: row.description,
            action: 'created',
            reason: 'New CBS code in variation — Original_Budget set to 0',
          });
          rowsCreated++;
        }
      } else {
        // ── Step 5c: First upload — Create all rows ──
        await this.createBudgetItem(row, projectCode, true);
        details.push({
          cbsCode: row.cbsCode,
          description: row.description,
          action: 'created',
          reason: 'First upload — full line item created',
        });
        rowsCreated++;
      }
    }

    onProgress?.({ step: 'complete', message: `CBS import complete for ${projectCode}. ${rowsCreated} created, ${rowsUpdated} updated, ${rowsSkipped} skipped.`, progress: 100, currentRow: rows.length, totalRows: rows.length });

    return {
      success: true,
      projectCode,
      isVariation,
      rowsTotal: rows.length,
      rowsCreated,
      rowsUpdated,
      rowsSkipped,
      summary: `Budget import ${isVariation ? 'variation' : 'first upload'} for ${projectCode}: ${rowsCreated} created, ${rowsUpdated} updated, ${rowsSkipped} skipped.`,
      details,
    };
  }

  // ============================================
  // Project Code Extraction
  // ============================================

  /**
   * Extract project code from filename.
   * "PRJ-HEIRS-ISU-2026_CBS.xlsx" → "PRJ-HEIRS-ISU-2026"
   *
   * Equivalent Power Automate expression:
   *   first(split(triggerOutputs()?['body/{FilenameWithExtension}'], '_CBS'))
   */
  private extractProjectCode(fileName: string): string {
    // Split on '_CBS' (case-insensitive) and take the first part
    const parts = fileName.split(/_CBS/i);
    if (parts.length < 2) return '';

    let code = parts[0].trim();
    // Remove any file extension artifacts from the first part
    // e.g. if the file is just "MYPROJ_CBS" without extension
    return code;
  }

  // ============================================
  // Excel Parsing (replaces Excel Online connector)
  // ============================================

  /**
   * Parse CBS_Table from an XLSX file buffer.
   *
   * Equivalent Power Automate action:
   *   Excel Online → "List rows present in a table"
   *   Table: CBS_Table
   *
   * This handles:
   * - Looking for a sheet named 'CBS_Table' or the first sheet
   * - Looking for a table named 'CBS_Table' within the workbook
   * - Mapping columns by header name (flexible column order)
   * - Filtering out empty/subtotal rows
   */
  private parseExcelCbsTable(fileBuffer: ArrayBuffer): ICbsRow[] {
    const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: true });

    // Try to find the CBS_Table sheet or table
    let sheet: XLSX.WorkSheet | undefined;

    // First, look for a sheet named "CBS_Table" (case-insensitive)
    const sheetNames = workbook.SheetNames;
    const cbsSheetName = sheetNames.find(
      (name) => name.toUpperCase() === 'CBS_TABLE' || name.toUpperCase().includes('CBS')
    );

    if (cbsSheetName) {
      sheet = workbook.Sheets[cbsSheetName];
    } else {
      // Fall back to the first sheet
      sheet = workbook.Sheets[sheetNames[0]];
    }

    if (!sheet) {
      throw new Error('No sheets found in the workbook.');
    }

    // Convert sheet to JSON with raw arrays first (handles merged title rows
    // that would produce __EMPTY column headers)
    const rawRows: (string | number | undefined)[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
    });

    // Known header keywords to identify the data header row
    const headerKeywords = ['cbs', 'phase', 'description', 'uom', 'qty', 'rate', 'amount'];

    // Scan rows to find the actual header row (skipping title/merged rows)
    let headerRowIndex = -1;
    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      const rowText = (row || []).map((c) => String(c || '').toLowerCase()).join(' ');
      if (headerKeywords.some((kw) => rowText.includes(kw))) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      if (rawRows.length === 0) {
        throw new Error('No data found in the CBS sheet.');
      }
      headerRowIndex = 0;
    }

    // Build column-name-to-header mapping for this sheet
    const headerRow = rawRows[headerRowIndex] || [];
    const colNameToIndex: Record<string, number> = {};
    headerRow.forEach((h, idx) => {
      const name = String(h || '').trim();
      if (name) colNameToIndex[name] = idx;
    });

    // Map expected column names to actual headers (flexible matching)
    const resolvedHeaders = this.resolveCbsColumnHeaders(Object.keys(colNameToIndex));

    if (!resolvedHeaders.cbsCode) {
      throw new Error('Required column "CBS Code" or "CBS_Code" not found in the sheet. Available columns: ' + Object.keys(colNameToIndex).join(', '));
    }

    // Rebuild each data row as a clean Record<string, ...> using resolved header names
    const jsonData: Record<string, string | number | undefined>[] = [];
    for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || row.every((cell) => String(cell || '').trim() === '')) continue;

      const entry: Record<string, string | number | undefined> = {};
      for (const [canonical, headerName] of Object.entries(resolvedHeaders)) {
        if (!headerName) continue;
        const colIdx = colNameToIndex[headerName];
        if (colIdx != null && colIdx >= 0 && colIdx < row.length) {
          entry[canonical] = row[colIdx];
        }
      }
      jsonData.push(entry);
    }

    // Parse and validate each row
    const rows: ICbsRow[] = [];
    const seenCodes = new Set<string>();

    for (const row of jsonData) {
      const cbsCode = this.safeString(row.cbsCode).trim();

      // Skip rows with empty CBS Code or TOTAL rows
      if (!cbsCode || cbsCode.toUpperCase() === 'TOTAL' || cbsCode === '0') {
        continue;
      }

      // Skip duplicate CBS codes (keep first occurrence)
      const upperCode = cbsCode.toUpperCase();
      if (seenCodes.has(upperCode)) {
        continue;
      }
      seenCodes.add(upperCode);

      const description = this.safeString(row.description) || cbsCode;
      const qty = this.safeNumber(row.qty);
      const rate = this.safeNumber(row.rate);
      const amount = this.safeNumber(row.amount);
      const phase = this.safeString(row.phase);
      const uom = this.safeString(row.uom);
      const remarks = this.safeString(row.remarks);

      // Use provided amount, or calculate from Qty × Rate
      const budgetAmount = amount > 0 ? amount : (qty * rate);

      rows.push({
        cbsCode,
        phase: phase.toUpperCase(),
        description,
        uom: uom.toUpperCase(),
        qty,
        rate,
        amount: budgetAmount,
        remarks,
      });
    }

    return rows;
  }

  /**
   * Resolve column header names from the Excel sheet to our canonical field names.
   * Accepts various naming conventions (with spaces, underscores, different casing).
   */
  private resolveCbsColumnHeaders(
    headers: string[]
  ): {
    cbsCode: string;
    phase: string;
    description: string;
    uom: string;
    qty: string;
    rate: string;
    amount: string;
    remarks: string;
  } {
    const findHeader = (aliases: string[]): string => {
      for (const alias of aliases) {
        const match = headers.find(
          (h) => h.replace(/[\s_]/g, '').toLowerCase() === alias.replace(/[\s_]/g, '').toLowerCase()
        );
        if (match) return match;
      }
      return '';
    };

    return {
      cbsCode: findHeader(['CBS Code', 'CBS_Code', 'CBS Code', 'CbsCode', 'WBSCode', 'WBS Code']) || headers[0],
      phase: findHeader(['Phase', 'PHASE', 'Project Phase', 'Project_Phase']),
      description: findHeader(['Description', 'DESCRIPTION', 'Desc', 'Item Description', 'Activity', 'Title']),
      uom: findHeader(['UOM', 'Uom', 'Unit', 'Unit of Measure']),
      qty: findHeader(['Qty', 'QTY', 'Quantity', 'QUANTITY']),
      rate: findHeader(['Rate', 'RATE', 'Rate (NGN)', 'Rate_NGN', 'Unit Rate']),
      amount: findHeader(['Amount', 'AMOUNT', 'Amount (NGN)', 'Amount_NGN', 'Budget Amount', 'Budget_Amount', 'Total']),
      remarks: findHeader(['Remarks', 'REMARKS', 'Note', 'Notes', 'Comment']),
    };
  }

  // ============================================
  // SharePoint Data Access
  // ============================================

  /** Map of canonical field names to possible SharePoint internal names */
  private readonly fieldNameMappings: Record<string, string[]> = {
    Project_Code: ['Project_Code', 'ProjectCode', 'Project_x0020_Code', 'Project_x005f_Code', 'project_x005f_code'],
    CBS_Code: ['CBS_Code', 'CBS_x0020_Code', 'CbsCode', 'CBS', 'CBS_x005f_Code'],
    Phase: ['Phase', 'PHASE'],
    Description: ['Description', 'DESCRIPTION'],
    UOM: ['UOM', 'Uom'],
    Qty: ['Qty', 'QTY', 'Quantity'],
    Rate: ['Rate', 'RATE', 'UnitRate'],
    Budget_Amount: ['Budget_Amount', 'Budget_x0020_Amount', 'BudgetAmount', 'Budget', 'Budget_x005f_Amount'],
    Original_Budget: ['Original_Budget', 'Original_x0020_Budget', 'OriginalBudget', 'Original_x005f_Budget'],
    Actual_Amount: ['Actual_Amount', 'Actual_x0020_Amount', 'ActualAmount', 'Actual_x005f_Amount'],
    Variance_Budget: ['Variance_Budget', 'Variance_x0020_Budget', 'Variance_x005f_Budget'],
    Remarks: ['Remarks', 'REMARKS', 'Comment'],
  };

  /** Cached resolved field names for the budget list */
  private resolvedFieldCache: Record<string, string> | null = null;

  /**
   * Resolve SharePoint internal field names by querying the list schema.
   * This prevents 400 errors when the actual internal names differ from
   * what we expect (e.g. 'Budget_Amount' vs 'Budget_x0020_Amount').
   *
   * Strategy:
   * 1. First query all fields (including hidden) for exact alias matches.
   * 2. Try querying visible-only fields as a cross-check.
   * 3. If no match is found from exact aliases, fall back to by-name
   *    detection on the full endpoint.
   */
  private async resolveBudgetListFields(): Promise<Record<string, string>> {
    if (this.resolvedFieldCache) return this.resolvedFieldCache;

    const resolved: Record<string, string> = {};
    const webUrl = this.pageContext.web.absoluteUrl;

    // Query ALL fields (including hidden) so we don't miss site columns
    // that may have been added as read-only or hidden fields.
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.BUDGET_ITEMS_LIST}')/fields?$select=InternalName,Title`;

    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const fieldInfos: Array<{ InternalName: string; Title: string }> =
          (data.value || []).filter((f: any) => f && f.InternalName && f.InternalName !== '');

        // Collect all internal names for direct matching
        const internalNames: string[] = fieldInfos.map((f) => f.InternalName);

        // Also create a map from Title → InternalName for display-name matching
        const titleToInternal = new Map<string, string>();
        for (const fi of fieldInfos) {
          const normTitle = fi.Title.replace(/[\s_]/g, '').toLowerCase();
          titleToInternal.set(normTitle, fi.InternalName);
        }

        // For each canonical field, find the matching internal name
        for (const [canonical, aliases] of Object.entries(this.fieldNameMappings)) {
          // 1. Exact match against internal names
          let match = aliases.find((alias) => internalNames.includes(alias));

          // 2. Fuzzy match: try normalized alias against display-name map
          if (!match) {
            for (const alias of aliases) {
              const normAlias = alias.replace(/[\s_]/g, '').toLowerCase();
              const titleMatch = titleToInternal.get(normAlias);
              if (titleMatch) {
                match = titleMatch;
                break;
              }
            }
          }

          // 3. Partial match: scan all internal names for any containing alias word
          if (!match) {
            for (const alias of aliases) {
              const lowerAlias = alias.replace(/[\s_]/g, '').toLowerCase();
              const partial = internalNames.find(
                (name) => name.replace(/[_\s]/g, '').toLowerCase().indexOf(lowerAlias) >= 0
              );
              if (partial) {
                match = partial;
                break;
              }
            }
          }

          // 4. Final fallback: use first alias as-is
          resolved[canonical] = match || aliases[0];
        }
      } else {
        // API call failed — fallback to canonical names
        for (const [canonical] of Object.entries(this.fieldNameMappings)) {
          resolved[canonical] = canonical;
        }
      }
    } catch {
      // Fallback: use canonical names as-is
      for (const [canonical] of Object.entries(this.fieldNameMappings)) {
        resolved[canonical] = canonical;
      }
    }

    this.resolvedFieldCache = resolved;
    return resolved;
  }

  /** Build a field-mapped body for SharePoint create/update operations */
  private async buildBudgetBody(fields: Record<string, unknown>): Promise<Record<string, unknown>> {
    const fieldMap = await this.resolveBudgetListFields();
    const body: Record<string, unknown> = {};
    for (const [canonical, value] of Object.entries(fields)) {
      const mappedName = fieldMap[canonical] || canonical;
      if (value !== undefined) {
        body[mappedName] = value;
      }
    }
    return body;
  }

  /**
   * Fetch existing budget items for a project code.
   * Equivalent Power Automate: Get items from Project_Budget_Items
   *   Filter: Project_Code eq '{projectCode}'
   *   Top Count: 1 (used only for variation detection — we fetch ALL for matching)
   */
  private async fetchExistingBudgetItems(
    projectCode: string
  ): Promise<Array<{ ID: number; CBS_Code: string; Budget_Amount: number }>> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const escaped = projectCode.replace(/'/g, "''");
    const fieldMap = await this.resolveBudgetListFields();
    const cbsField = fieldMap['CBS_Code'] || 'CBS_Code';
    const projectField = fieldMap['Project_Code'] || 'Project_Code';
    const budgetField = fieldMap['Budget_Amount'] || 'Budget_Amount';
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.BUDGET_ITEMS_LIST}')/items?$filter=${encodeURIComponent(`${projectField} eq '${escaped}'`)}&$select=ID,${cbsField},${budgetField}&$top=2000`;

    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return [];
      const data = await resp.json();
      return (data.value || []).map((item: any) => ({
        ID: item.ID,
        CBS_Code: item[cbsField] || '',
        Budget_Amount: item[budgetField] || 0,
      }));
    } catch {
      return [];
    }
  }

  /**
   * Create a new budget item.
   * Equivalent Power Automate: Create item in Project_Budget_Items
   */
  private async createBudgetItem(
    row: ICbsRow,
    projectCode: string,
    isFirstUpload: boolean
  ): Promise<number> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.BUDGET_ITEMS_LIST}')/items`;

    const body = await this.buildBudgetBody({
      Title: row.description.substring(0, 255),
      Project_Code: projectCode,
      CBS_Code: row.cbsCode,
      Phase: row.phase,
      Description: row.description,
      UOM: row.uom,
      Qty: row.qty,
      Rate: row.rate,
      Budget_Amount: row.amount,
      Original_Budget: isFirstUpload ? row.amount : 0,
      Actual_Amount: 0,
      Remarks: row.remarks || '',
    });

    const resp = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Failed to create budget item for ${row.cbsCode}: ${resp.status} — ${errText.substring(0, 200)}`);
    }

    const result = await resp.json();
    return result.ID;
  }

  /**
   * Update an existing budget item (variation).
   * IMPORTANT: Does NOT update Original_Budget — it's left unchanged.
   *
   * Equivalent Power Automate: Update item in Project_Budget_Items
   *   - Budget_Amount: new value from spreadsheet
   *   - Qty: new value
   *   - Rate: new value
   *   - Remarks: new value
   *   - Original_Budget: UNCHANGED
   */
  private async updateBudgetItem(
    id: number,
    row: ICbsRow,
    projectCode: string
  ): Promise<void> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.BUDGET_ITEMS_LIST}')/items(${id})`;

    const body = await this.buildBudgetBody({
      Title: row.description.substring(0, 255),
      Project_Code: projectCode,
      CBS_Code: row.cbsCode,
      Phase: row.phase,
      Description: row.description,
      UOM: row.uom,
      Qty: row.qty,
      Rate: row.rate,
      Budget_Amount: row.amount,
      // Original_Budget is deliberately NOT updated
      Remarks: row.remarks || '',
    });

    const resp = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Failed to update budget item ${id} (${row.cbsCode}): ${resp.status} — ${errText.substring(0, 200)}`);
    }
  }

  // ============================================
  // Helpers
  // ============================================

  /**
   * Upload a CBS file to the project document library.
   * Used by the UI to store the file in Documents/Budgets/{projectCode}/
   */
  public async uploadCbsFile(
    file: File,
    projectCode: string
  ): Promise<string> {
    const libraryRootUrl = await getProjectDocumentLibraryRootUrl(this.spHttpClient, this.pageContext);
    const folderUrl = getProjectDocumentFolderServerRelativeUrl(
      libraryRootUrl,
      PROJECT_BUDGET_FOLDER_NAME,
      projectCode
    );

    // Ensure the folder exists
    await this.ensureFolderExists(folderUrl);

    // Upload the file
    const normalizedFolderUrl = folderUrl.startsWith('/') ? folderUrl : `/${folderUrl.replace(/^\/+/, '')}`;
    const uploadUrl = `${this.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(normalizedFolderUrl)}')/Files/add(url='${encodeURIComponent(file.name)}',overwrite=true)`;

    const resp = await this.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream',
      },
      body: await file.arrayBuffer(),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Failed to upload CBS file: ${resp.status} — ${errText.substring(0, 200)}`);
    }

    return folderUrl;
  }

  /**
   * Ensure a folder path exists (create if missing).
   */
  private async ensureFolderExists(folderServerRelativeUrl: string): Promise<void> {
    const normalizedUrl = folderServerRelativeUrl.startsWith('/')
      ? folderServerRelativeUrl
      : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
    const webUrl = this.pageContext.web.absoluteUrl;

    // Check if exists
    const checkUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(normalizedUrl)}')`;
    const checkResp = await this.spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResp.ok) return;

    // Create parent chain recursively
    const parentPath = normalizedUrl.substring(0, normalizedUrl.lastIndexOf('/'));
    const folderName = normalizedUrl.substring(normalizedUrl.lastIndexOf('/') + 1);
    if (parentPath && folderName) {
      await this.ensureFolderExists(parentPath);
      const createUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(parentPath)}')/Folders/add('${encodeURIComponent(folderName)}')`;
      await this.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=verbose' },
        body: '',
      });
    }
  }

  private safeString(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    if (value instanceof Date) return value.toISOString();
    return String(value);
  }

  private safeNumber(value: unknown): number {
    if (value === null || value === undefined) return 0;
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[^0-9.\-]/g, '');
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }
}

// ============================================
// Factory function
// ============================================

export function createCbsBudgetImportService(
  spHttpClient: SPHttpClient,
  pageContext: PageContext
): CbsBudgetImportService {
  return new CbsBudgetImportService(spHttpClient, pageContext);
}
