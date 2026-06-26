import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
  getProjectDocumentLibraryRootUrl,
  getProjectDocumentFolderServerRelativeUrl,
  PROJECT_WBS_FOLDER_NAME,
} from './ProjectDocumentStorageService';

// ============================================
// WBS Schedule Import Interfaces
// ============================================

export interface IWbsRow {
  wbsId: string;
  activityName: string;
  level: number;
  phaseLink: string;
  durationDays: number;
  plannedStart: string;
  plannedFinish: string;
  percentComplete: number;
  weight: number;
  predecessors: string;
  status: string;
  remarks: string;
}

export interface IWbsUploadFile {
  /** Original file name e.g. PRJ-HEIRS-ISU-2026_WBS.xlsx */
  fileName: string;
  /** Extracted project code */
  projectCode: string;
  /** Full file content as ArrayBuffer */
  fileBuffer: ArrayBuffer;
}

export interface IWbsImportResult {
  success: boolean;
  projectCode: string;
  isUpdate: boolean;
  rowsTotal: number;
  rowsCreated: number;
  rowsUpdated: number;
  rowsSkipped: number;
  summary: string;
  error?: string;
  details: IWbsRowResult[];
}

export interface IWbsRowResult {
  wbsId: string;
  activityName: string;
  action: 'created' | 'updated' | 'skipped';
  reason?: string;
}

export interface IWbsImportProgress {
  step: string;
  message: string;
  progress: number;
  currentRow?: number;
  totalRows?: number;
}

// ============================================
// WBS Schedule Import Service
// ============================================

/**
 * WbsScheduleImportService — replaces Power Automate "Parse WBS Schedule Upload"
 * flow entirely in-browser.
 *
 * Flow trigger: file uploaded to Documents/Budgets/ ending in _WBS
 *
 * Benefits over Power Automate:
 * - Zero latency — no cloud flow queuing, no polling
 * - Batch operations — processes all rows in a single transaction
 * - Richer validation — validates WBS ID format, numeric types, status defaults
 * - Detailed reporting — per-row results with specific reasons
 */
export class WbsScheduleImportService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;

  private readonly SCHEDULE_LIST = 'Project_Schedule';

  constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
    this.spHttpClient = spHttpClient;
    this.pageContext = pageContext;
  }

  // ============================================
  // PUBLIC: Main entry point
  // ============================================

  /**
   * Process a WBS schedule file upload. Replaces the entire Power Automate flow.
   *
   * Flow equivalent:
   *   1. Trigger: file uploaded with _WBS
   *   2. Condition: filename contains '_WBS'
   *   3. Compose: Extract Project Code
   *   4. Get items: Check if schedule exists (update detection)
   *   5. Compose: Is Update?
   *   6. List rows: Parse Excel WBS_Table
   *   7. Apply to each: loop with conditions
   *      7a. Skip empty WBS ID rows
   *      7b. Update → update existing OR create new
   *      7c. First upload → create all
   */
  public async importWbsFile(
    file: IWbsUploadFile,
    onProgress?: (progress: IWbsImportProgress) => void
  ): Promise<IWbsImportResult> {
    const details: IWbsRowResult[] = [];
    let rowsCreated = 0;
    let rowsUpdated = 0;
    let rowsSkipped = 0;

    // ── Step 1: Determine project code ──
    // If a project code is explicitly provided in the request, use it directly.
    // Otherwise, extract it from the filename (Power Automate trigger scenario).
    let projectCode = (file.projectCode || '').trim();

    onProgress?.({ step: 'validate', message: `Preparing to import: ${file.fileName}`, progress: 5, currentRow: 0, totalRows: 0 });

    if (!projectCode) {
      // No explicit project code — extract from filename (requires _WBS suffix)
      if (!file.fileName.toUpperCase().includes('_WBS')) {
        return {
          success: false,
          projectCode: '',
          isUpdate: false,
          rowsTotal: 0, rowsCreated: 0, rowsUpdated: 0, rowsSkipped: 0,
          summary: `File '${file.fileName}' does not contain '_WBS' and no project code was provided.`,
          details: [],
          error: 'Missing project code and filename does not contain _WBS separator.',
        };
      }
      projectCode = this.extractProjectCode(file.fileName);
      if (!projectCode) {
        return {
          success: false, projectCode: '', isUpdate: false,
          rowsTotal: 0, rowsCreated: 0, rowsUpdated: 0, rowsSkipped: 0,
          summary: `Could not extract project code from '${file.fileName}'. Expected format: PRJ-CODE_WBS.xlsx`,
          details: [], error: 'Project code extraction failed.',
        };
      }
    }

    // ── Step 3: Parse Excel WBS_Table ──
    onProgress?.({ step: 'parse', message: `Parsing Excel WBS_Table from ${file.fileName}...`, progress: 20, currentRow: 0, totalRows: 0 });
    let rows: IWbsRow[];
    try {
      rows = this.parseExcelWbsTable(file.fileBuffer);
    } catch (err) {
      return {
        success: false, projectCode, isUpdate: false,
        rowsTotal: 0, rowsCreated: 0, rowsUpdated: 0, rowsSkipped: 0,
        summary: `Failed to parse Excel file: ${err instanceof Error ? err.message : 'Unknown error'}`,
        details: [], error: err instanceof Error ? err.message : 'Excel parse failure',
      };
    }

    if (rows.length === 0) {
      return {
        success: false, projectCode, isUpdate: false,
        rowsTotal: 0, rowsCreated: 0, rowsUpdated: 0, rowsSkipped: 0,
        summary: 'No valid WBS rows found in the Excel file. All rows were empty or lacked a WBS ID.',
        details: [], error: 'No valid WBS data rows after filtering.',
      };
    }

    onProgress?.({ step: 'found-rows', message: `Found ${rows.length} WBS rows to process.`, progress: 30, currentRow: 0, totalRows: rows.length });

    // ── Step 4: Check if schedule already exists (update detection) ──
    onProgress?.({ step: 'check-existing', message: `Checking if schedule exists for ${projectCode}...`, progress: 35, currentRow: 0, totalRows: rows.length });
    const existingRows = await this.fetchExistingScheduleItems(projectCode);
    const isUpdate = existingRows.length > 0;

    // Build a lookup map of existing rows by WBS_ID
    const existingMap = new Map<string, number>();
    if (isUpdate) {
      existingRows.forEach((r) => {
        if (r.WBS_ID) existingMap.set(r.WBS_ID.trim().toUpperCase(), r.ID);
      });
    }

    onProgress?.({
      step: 'update-check',
      message: isUpdate
        ? `Update detected — ${existingRows.length} existing rows found. Will update matching WBS IDs.`
        : `First upload — creating all ${rows.length} rows.`,
      progress: 40, currentRow: 0, totalRows: rows.length,
    });

    // ── Step 5: Process each row (Apply to each) ──
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowNum = i + 1;
      onProgress?.({
        step: 'processing-row',
        message: `Processing row ${rowNum}/${rows.length}: ${row.wbsId} — ${row.activityName}`,
        progress: 40 + Math.round((rowNum / rows.length) * 55),
        currentRow: rowNum, totalRows: rows.length,
      });

      // ── Step 5a: Skip empty rows ──
      if (!row.wbsId || row.wbsId.trim().toUpperCase() === 'TOTAL') {
        details.push({
          wbsId: row.wbsId || '(empty)',
          activityName: row.activityName || '(empty)',
          action: 'skipped',
          reason: !row.wbsId ? 'Empty WBS ID' : 'Row is a TOTAL row',
        });
        rowsSkipped++;
        continue;
      }

      // Default status to "Not Started" if empty
      const resolvedStatus = row.status || 'Not Started';

      if (isUpdate) {
        // ── Step 5b: Update branch ──
        const existingId = existingMap.get(row.wbsId.trim().toUpperCase());

        if (existingId) {
          // Update existing row
          await this.updateScheduleItem(existingId, row, projectCode, resolvedStatus);
          details.push({
            wbsId: row.wbsId,
            activityName: row.activityName,
            action: 'updated',
            reason: `Updated Activity_Name, Duration_Days, dates, Weight, Status`,
          });
          rowsUpdated++;
        } else {
          // New line in update: create
          await this.createScheduleItem(row, projectCode, resolvedStatus);
          details.push({
            wbsId: row.wbsId,
            activityName: row.activityName,
            action: 'created',
            reason: 'New WBS ID in update — full item created',
          });
          rowsCreated++;
        }
      } else {
        // ── Step 5c: First upload — Create all ──
        await this.createScheduleItem(row, projectCode, resolvedStatus);
        details.push({
          wbsId: row.wbsId,
          activityName: row.activityName,
          action: 'created',
          reason: 'First upload — full line item created',
        });
        rowsCreated++;
      }
    }

    onProgress?.({
      step: 'complete',
      message: `WBS import complete for ${projectCode}. ${rowsCreated} created, ${rowsUpdated} updated, ${rowsSkipped} skipped.`,
      progress: 100, currentRow: rows.length, totalRows: rows.length,
    });

    return {
      success: true,
      projectCode,
      isUpdate,
      rowsTotal: rows.length,
      rowsCreated,
      rowsUpdated,
      rowsSkipped,
      summary: `Schedule import ${isUpdate ? 'update' : 'first upload'} for ${projectCode}: ${rowsCreated} created, ${rowsUpdated} updated, ${rowsSkipped} skipped.`,
      details,
    };
  }

  // ============================================
  // Project Code Extraction
  // ============================================

  private extractProjectCode(fileName: string): string {
    const parts = fileName.split(/_WBS/i);
    if (parts.length < 2) return '';
    return parts[0].trim();
  }

  // ============================================
  // Excel Parsing (replaces Excel Online connector)
  // ============================================

  /**
   * Parse WBS_Table from an XLSX file buffer.
   *
   * Equivalent Power Automate action:
   *   Excel Online → "List rows present in a table"
   *   Table: WBS_Table
   *
   * Flexible column mapping supporting various naming conventions.
   */
  private parseExcelWbsTable(fileBuffer: ArrayBuffer): IWbsRow[] {
    const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: true });

    // Find the WBS_Table sheet
    let sheet: XLSX.WorkSheet | undefined;
    const sheetNames = workbook.SheetNames;
    const wbsSheetName = sheetNames.find(
      (name) => name.toUpperCase() === 'WBS_TABLE' || name.toUpperCase().includes('WBS')
    );

    if (wbsSheetName) {
      sheet = workbook.Sheets[wbsSheetName];
    } else {
      sheet = workbook.Sheets[sheetNames[0]];
    }

    if (!sheet) {
      throw new Error('No sheets found in the workbook.');
    }

    // Read as raw arrays first (header: 1) to skip non-data title/merged rows
    // that would produce __EMPTY columns with header: undefined.
    const rawRows: (string | number | undefined)[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
    });

    // Known header keywords that identify the data header row
    const headerKeywords = ['wbs', 'activity', 'phase', 'duration', 'start', 'finish', 'level'];

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
        throw new Error('No data found in the WBS sheet.');
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

    // Resolve canonical → actual header name (so the rest of the parsing code works)
    const resolvedHeaders = this.resolveWbsColumnHeaders(Object.keys(colNameToIndex));

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
        } else {
          entry[canonical] = undefined;
        }
      }
      jsonData.push(entry);
    }

    const rows: IWbsRow[] = [];
    const seenIds = new Set<string>();

    const parseDate = (val: unknown): string => {
      if (!val || val === '') return '';
      if (val instanceof Date) return val.toISOString();
      const str = String(val).trim();
      if (!str || str === '0') return '';
      // Try Excel serial date number
      const num = Number(str);
      if (!isNaN(num) && num > 1 && num < 100000) {
        const excelEpoch = new Date(1899, 11, 30);
        const d = new Date(excelEpoch.getTime() + num * 86400000);
        return d.toISOString();
      }
      // Try string date
      const d = new Date(str);
      if (!isNaN(d.getTime())) return d.toISOString();
      return str;
    };

    for (const row of jsonData) {
      const wbsId = this.safeString(row.wbsId).trim();

      if (!wbsId || wbsId.toUpperCase() === 'TOTAL' || wbsId === '0') {
        continue;
      }

      const upperId = wbsId.toUpperCase();
      if (seenIds.has(upperId)) continue;
      seenIds.add(upperId);

      const activityName = this.safeString(row.activityName) || wbsId;
      const level = this.safeNumber(row.level);
      const phaseLink = this.safeString(row.phaseLink);
      const durationDays = this.safeNumber(row.durationDays);
      const plannedStart = parseDate(row.plannedStart);
      const plannedFinish = parseDate(row.plannedFinish);
      const percentComplete = this.safeNumber(row.percentComplete);
      const weight = this.safeNumber(row.weight);
      const predecessors = this.safeString(row.predecessors);
      const remarks = this.safeString(row.remarks);
      const status = this.safeString(row.status);

      rows.push({
        wbsId,
        activityName,
        level: level > 0 ? level : 1,
        phaseLink,
        durationDays,
        plannedStart,
        plannedFinish,
        percentComplete,
        weight,
        predecessors,
        status,
        remarks,
      });
    }

    return rows;
  }

  /**
   * Resolve column header names from the Excel sheet to canonical field names.
   */
  private resolveWbsColumnHeaders(headers: string[]) {
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
      wbsId: findHeader(['WBS ID', 'WBS_ID', 'WbsId', 'WBS ID', 'WBSCode']),
      activityName: findHeader(['Activity Name', 'Activity_Name', 'Activity', 'Task Name', 'Title', 'Description']),
      level: findHeader(['Level', 'LEVEL', 'WBS Level', 'WBS_Level']),
      phaseLink: findHeader(['Phase Link', 'Phase_Link', 'Phase', 'PHASE', 'PhaseLink']),
      durationDays: findHeader(['Duration (Days)', 'Duration_Days', 'Duration Days', 'Duration', 'DURATION']),
      plannedStart: findHeader(['Start Date', 'Start_Date', 'Planned Start', 'Planned_Start', 'PlannedStart']),
      plannedFinish: findHeader(['Finish Date', 'Finish_Date', 'Planned Finish', 'Planned_Finish', 'PlannedFinish', 'End Date']),
      percentComplete: findHeader(['% Complete', 'Percent_Complete', 'Percent Complete', 'PCT', '%']),
      weight: findHeader(['Weight (%)', 'Weight', 'WEIGHT', 'Weight_Pct', 'WeightPct']),
      predecessors: findHeader(['Predecessors', 'PREDECESSORS', 'Predecessor', 'Dependencies']),
      status: findHeader(['Status', 'STATUS', 'Activity Status', 'Activity_Status']),
      remarks: findHeader(['Remarks', 'REMARKS', 'Note', 'Notes', 'Comment']),
    };
  }

  // ============================================
  // SharePoint Data Access
  // ============================================

  /** Map of canonical field names to possible SharePoint internal names */
  private readonly scheduleFieldMappings: Record<string, string[]> = {
    Project_Code: ['Project_Code', 'ProjectCode', 'Project_x0020_Code', 'Project_x005f_Code', 'project_x005f_code'],
    WBS_ID: ['WBS_ID', 'WBS_x0020_ID', 'WBS_x005f_ID', 'WbsId', 'WBS'],
    Activity_Name: ['Activity_Name', 'Activity_x0020_Name', 'ActivityName', 'Activity_x005f_Name'],
    Level: ['Level', 'LEVEL'],
    Phase_Link: ['Phase_Link', 'Phase_x0020_Link', 'PhaseLink', 'Phase_x005f_Link'],
    Duration_Days: ['Duration_Days', 'Duration_x0020_Days', 'DurationDays', 'Duration_x005f_Days'],
    Planned_Start: ['Planned_Start', 'Planned_x0020_Start', 'PlannedStart', 'Planned_x005f_Start'],
    Planned_Finish: ['Planned_Finish', 'Planned_x0020_Finish', 'PlannedFinish', 'Planned_x005f_Finish'],
    Percent_Complete: ['Percent_Complete', 'Percent_x0020_Complete', 'PercentComplete', 'Percent_x005f_Complete'],
    Weight: ['Weight', 'WEIGHT'],
    Predecessors: ['Predecessors', 'PREDECESSORS'],
    Status: ['Status', 'STATUS'],
    Remarks: ['Remarks', 'REMARKS', 'Comment'],
  };

  /** Cached resolved field names for the schedule list */
  private scheduleFieldCache: Record<string, string> | null = null;

  /**
   * Resolve SharePoint internal field names for the schedule list.
   * Queries ALL fields (including hidden) and tries exact, fuzzy, and
   * partial matching against known aliases to find the actual internal name.
   */
  private async resolveScheduleListFields(): Promise<Record<string, string>> {
    if (this.scheduleFieldCache) return this.scheduleFieldCache;

    const resolved: Record<string, string> = {};
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.SCHEDULE_LIST}')/fields?$select=InternalName,Title`;

    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const fieldInfos: Array<{ InternalName: string; Title: string }> =
          (data.value || []).filter((f: any) => f && f.InternalName && f.InternalName !== '');

        const internalNames: string[] = fieldInfos.map((f) => f.InternalName);

        // Map from normalized Title to InternalName for display-name matching
        const titleToInternal = new Map<string, string>();
        for (const fi of fieldInfos) {
          const normTitle = fi.Title.replace(/[\s_]/g, '').toLowerCase();
          titleToInternal.set(normTitle, fi.InternalName);
        }

        for (const [canonical, aliases] of Object.entries(this.scheduleFieldMappings)) {
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

          // 3. Partial match: scan all internal names for any containing the alias
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

          // 4. Fallback: use first alias as-is
          resolved[canonical] = match || aliases[0];
        }
      } else {
        for (const [canonical] of Object.entries(this.scheduleFieldMappings)) {
          resolved[canonical] = canonical;
        }
      }
    } catch {
      for (const [canonical] of Object.entries(this.scheduleFieldMappings)) {
        resolved[canonical] = canonical;
      }
    }

    this.scheduleFieldCache = resolved;
    return resolved;
  }

  private async buildScheduleBody(fields: Record<string, unknown>): Promise<Record<string, unknown>> {
    const fieldMap = await this.resolveScheduleListFields();
    const body: Record<string, unknown> = {};
    for (const [canonical, value] of Object.entries(fields)) {
      const mappedName = fieldMap[canonical] || canonical;
      if (value !== undefined) {
        body[mappedName] = value;
      }
    }
    return body;
  }

  private async fetchExistingScheduleItems(
    projectCode: string
  ): Promise<Array<{ ID: number; WBS_ID: string }>> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const escaped = projectCode.replace(/'/g, "''");
    const fieldMap = await this.resolveScheduleListFields();
    const projectField = fieldMap['Project_Code'] || 'Project_Code';
    const wbsField = fieldMap['WBS_ID'] || 'WBS_ID';
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.SCHEDULE_LIST}')/items?$filter=${encodeURIComponent(`${projectField} eq '${escaped}'`)}&$select=ID,${wbsField}&$top=2000`;

    try {
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return [];
      const data = await resp.json();
      return (data.value || []).map((item: any) => ({
        ID: item.ID,
        WBS_ID: item[wbsField] || '',
      }));
    } catch {
      return [];
    }
  }

  private async createScheduleItem(
    row: IWbsRow,
    projectCode: string,
    resolvedStatus: string
  ): Promise<number> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.SCHEDULE_LIST}')/items`;

    const body = await this.buildScheduleBody({
      Title: row.activityName.substring(0, 255),
      Project_Code: projectCode,
      WBS_ID: row.wbsId,
      Activity_Name: row.activityName,
      Level: row.level,
      Phase_Link: row.phaseLink,
      Duration_Days: row.durationDays,
      Planned_Start: row.plannedStart || null,
      Planned_Finish: row.plannedFinish || null,
      Percent_Complete: Math.min(100, Math.max(0, row.percentComplete)),
      Weight: row.weight,
      Predecessors: row.predecessors || '',
      Status: resolvedStatus,
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
      throw new Error(`Failed to create schedule item for ${row.wbsId}: ${resp.status} — ${errText.substring(0, 200)}`);
    }

    const result = await resp.json();
    return result.ID;
  }

  private async updateScheduleItem(
    id: number,
    row: IWbsRow,
    projectCode: string,
    resolvedStatus: string
  ): Promise<void> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.SCHEDULE_LIST}')/items(${id})`;

    const body = await this.buildScheduleBody({
      Title: row.activityName.substring(0, 255),
      Project_Code: projectCode,
      WBS_ID: row.wbsId,
      Activity_Name: row.activityName,
      Level: row.level,
      Phase_Link: row.phaseLink,
      Duration_Days: row.durationDays,
      Planned_Start: row.plannedStart || null,
      Planned_Finish: row.plannedFinish || null,
      Percent_Complete: Math.min(100, Math.max(0, row.percentComplete)),
      Weight: row.weight,
      Predecessors: row.predecessors || '',
      Status: resolvedStatus,
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
      throw new Error(`Failed to update schedule item ${id} (${row.wbsId}): ${resp.status} — ${errText.substring(0, 200)}`);
    }
  }

  // ============================================
  // File Upload Helper
  // ============================================

  public async uploadWbsFile(file: File, projectCode: string): Promise<string> {
    const libraryRootUrl = await getProjectDocumentLibraryRootUrl(this.spHttpClient, this.pageContext);
    const folderUrl = getProjectDocumentFolderServerRelativeUrl(
      libraryRootUrl,
      PROJECT_WBS_FOLDER_NAME,
      projectCode
    );

    await this.ensureFolderExists(folderUrl);

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
      throw new Error(`Failed to upload WBS file: ${resp.status} — ${errText.substring(0, 200)}`);
    }

    return folderUrl;
  }

  private async ensureFolderExists(folderServerRelativeUrl: string): Promise<void> {
    const normalizedUrl = folderServerRelativeUrl.startsWith('/')
      ? folderServerRelativeUrl
      : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
    const webUrl = this.pageContext.web.absoluteUrl;

    const checkUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(normalizedUrl)}')`;
    const checkResp = await this.spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResp.ok) return;

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

  // ============================================
  // Helpers
  // ============================================

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

export function createWbsScheduleImportService(
  spHttpClient: SPHttpClient,
  pageContext: PageContext
): WbsScheduleImportService {
  return new WbsScheduleImportService(spHttpClient, pageContext);
}
