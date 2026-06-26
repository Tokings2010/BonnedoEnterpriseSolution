import * as React from 'react';
import {
  mergeStyleSets,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  Icon,
  CommandBar,
  ICommandBarItemProps,
  ProgressIndicator,
  PrimaryButton,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
  getProjectDocumentDisplayPath,
  getProjectDocumentFileName,
  getProjectDocumentLibraryRootUrl,
  getStoredProjectDocument,
  IStoredProjectDocument,
  PROJECT_WBS_FOLDER_NAME,
} from '../../services/ProjectDocumentStorageService';
import { createWbsScheduleImportService, IWbsImportProgress, IWbsImportResult } from '../../services/WbsScheduleImportService';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IProjectScheduleTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  projectCode: string;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IScheduleItem {
  ID: number;
  Project_Code: string;
  WBS_ID: string;
  Activity_Name: string;
  Level: number;
  Phase_Link: string;
  Duration_Days: number;
  Planned_Start: string;
  Planned_Finish: string;
  Actual_Start: string;
  Actual_Finish: string;
  Percent_Complete: number;
  Weight: number;
  Predecessors: string;
  Status: string;
  Remarks: string;
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const getStyles = () =>
  mergeStyleSets({
    container: {
      display: 'flex',
      flexDirection: 'column' as const,
      gap: '16px',
      overflow: 'auto',
      height: '100%',
    },
    summaryRow: {
      display: 'flex',
      gap: '16px',
      flexWrap: 'wrap' as const,
    },
    summaryCard: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px 24px',
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      minWidth: '150px',
      flex: '1',
    },
    cardValue: {
      fontSize: '22px',
      fontWeight: '700',
      color: '#1E2532',
    },
    cardLabel: {
      fontSize: '12px',
      color: '#6B7280',
      fontWeight: '500',
    },
    statusPill: {
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      display: 'inline-block',
    },
    progressCell: {
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      minWidth: '120px',
    },
    delayBadge: {
      padding: '2px 8px',
      borderRadius: '4px',
      fontSize: '10px',
      fontWeight: '600',
      backgroundColor: '#FEE2E2',
      color: '#EF4444',
    },
    overallProgress: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px 24px',
    },
    sectionTitle: {
      fontSize: '16px',
      fontWeight: '600',
      color: '#1E2532',
      marginTop: '8px',
    },
  });

// ─── Helper Functions ────────────────────────────────────────────────────────

const getStatusStyle = (status: string): { color: string; bg: string } => {
  switch (status) {
    case 'Completed': return { color: '#FFFFFF', bg: '#10B981' };
    case 'In Progress': return { color: '#FFFFFF', bg: '#2563EB' };
    case 'Delayed': return { color: '#FFFFFF', bg: '#EF4444' };
    case 'On Hold': return { color: '#FFFFFF', bg: '#F59E0B' };
    case 'Not Started':
    default: return { color: '#6B7280', bg: '#E5E7EB' };
  }
};

const formatDate = (dateStr: string): string => {
  if (!dateStr) return '—';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

const getDaysVariance = (planned: string, actual: string): number | null => {
  if (!planned) return null;
  const pDate = new Date(planned);
  const aDate = actual ? new Date(actual) : new Date();
  if (isNaN(pDate.getTime())) return null;
  return Math.round((aDate.getTime() - pDate.getTime()) / (1000 * 60 * 60 * 24));
};

// ─── Component ───────────────────────────────────────────────────────────────

const ProjectScheduleTab: React.FC<IProjectScheduleTabProps> = ({
  spHttpClient,
  pageContext,
  projectCode,
  isMobileView,
  onRefresh,
}) => {
  const styles = getStyles();
  const [items, setItems] = React.useState<IScheduleItem[]>([]);
  const [storedDocument, setStoredDocument] = React.useState<IStoredProjectDocument | undefined>(undefined);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [importResult, setImportResult] = React.useState<IWbsImportResult | null>(null);
  const [importProgress, setImportProgress] = React.useState<IWbsImportProgress | null>(null);
  const [isImporting, setIsImporting] = React.useState(false);
  const [importError, setImportError] = React.useState<string | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const wbsImportService = React.useMemo(
    () => createWbsScheduleImportService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  // ─── Fetch Data ──────────────────────────────────────────────────────────

  const fetchScheduleItems = React.useCallback(async (): Promise<void> => {
    if (!projectCode) {
      setItems([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const escapedProjectCode = projectCode.replace(/'/g, "''");
      const filteredUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$filter=Project_Code eq '${escapedProjectCode}'&$orderby=WBS_ID asc&$top=500`;
      const fallbackUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$top=5000`;
      let response: SPHttpClientResponse = await spHttpClient.get(
        filteredUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const fallbackResponse: SPHttpClientResponse = await spHttpClient.get(
          fallbackUrl,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (fallbackResponse.ok) {
          response = fallbackResponse;
        }
      }

      if (!response.ok) throw new Error(`Failed to fetch schedule: ${response.status}`);

      const data = await response.json();
      const today = new Date();

      const fetched: IScheduleItem[] = (data.value || [])
        .filter((item: any) => String(item.Project_Code || '').trim() === projectCode.trim())
        .map((item: any) => {
        let status = item.Status || 'Not Started';

        // Auto-detect delays
        if (item.Planned_Finish && status !== 'Completed' && status !== 'On Hold') {
          const plannedFinish = new Date(item.Planned_Finish);
          if (plannedFinish < today) {
            status = 'Delayed';
          }
        }

        return {
          ID: item.ID,
          Project_Code: item.Project_Code || '',
          WBS_ID: item.WBS_ID || '',
          Activity_Name: item.Activity_Name || item.Title || '',
          Level: item.Level || 1,
          Phase_Link: item.Phase_Link || '',
          Duration_Days: item.Duration_Days || 0,
          Planned_Start: item.Planned_Start || '',
          Planned_Finish: item.Planned_Finish || '',
          Actual_Start: item.Actual_Start || '',
          Actual_Finish: item.Actual_Finish || '',
          Percent_Complete: item.Percent_Complete || 0,
          Weight: item.Weight || 0,
          Predecessors: item.Predecessors || '',
          Status: status,
          Remarks: item.Remarks || '',
        };
      });

      setItems(fetched);
    } catch (err) {
      console.error('Error fetching schedule:', err);
      setError(err instanceof Error ? err.message : 'Failed to load schedule');
    } finally {
      setLoading(false);
    }
  }, [spHttpClient, pageContext, projectCode]);

  const fetchStoredDocument = React.useCallback(async (): Promise<void> => {
    if (!projectCode) {
      setStoredDocument(undefined);
      return;
    }

    try {
      const libraryRootUrl = await getProjectDocumentLibraryRootUrl(spHttpClient, pageContext);
      const document = await getStoredProjectDocument(
        spHttpClient,
        pageContext,
        libraryRootUrl,
        PROJECT_WBS_FOLDER_NAME,
        projectCode,
        'schedule'
      );

      setStoredDocument(document);
    } catch (error) {
      console.warn('[ProjectScheduleTab] Unable to read stored WBS source file:', error);
      setStoredDocument(undefined);
    }
  }, [spHttpClient, pageContext, projectCode]);

  React.useEffect(() => {
    fetchScheduleItems().catch((error) => console.error('Error refreshing schedule items:', error));
    fetchStoredDocument().catch((error) => console.error('Error refreshing stored schedule document:', error));
  }, [fetchScheduleItems, fetchStoredDocument]);

  // ─── WBS File Import Handler ───────────────────────────────────────────────

  const handleWbsFileImport = React.useCallback(async (file: File): Promise<void> => {
    setImportError(null);
    setImportResult(null);
    setIsImporting(true);
    setImportProgress({ step: 'start', message: 'Reading WBS file...', progress: 0, currentRow: 0, totalRows: 0 });

    try {
      const buffer = await file.arrayBuffer();
      const result = await wbsImportService.importWbsFile({
        fileName: file.name,
        projectCode,
        fileBuffer: buffer,
      }, (progress: IWbsImportProgress) => {
        setImportProgress({ ...progress });
      });

      setImportResult(result);
      if (result.success) {
        await fetchScheduleItems();
        await fetchStoredDocument();
        onRefresh();
      }
    } catch (err) {
      setImportError(err instanceof Error ? err.message : 'Failed to import WBS file');
    } finally {
      setIsImporting(false);
      setImportProgress(null);
    }
  }, [wbsImportService, projectCode, fetchScheduleItems, fetchStoredDocument, onRefresh]);

  const handleImportButtonClick = (): void => {
    fileInputRef.current?.click();
  };

  const handleFileSelected = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (file) {
      handleWbsFileImport(file).catch(console.error);
    }
    if (e.target) e.target.value = '';
  };

  // ─── Calculations ────────────────────────────────────────────────────────

  const totalActivities = items.length;
  const completed = items.filter((i) => i.Status === 'Completed').length;
  const inProgress = items.filter((i) => i.Status === 'In Progress').length;
  const delayed = items.filter((i) => i.Status === 'Delayed').length;

  // Weighted overall % complete
  const totalWeight = items.reduce((s, i) => s + i.Weight, 0);
  const weightedComplete = totalWeight > 0
    ? Math.round(items.reduce((s, i) => s + (i.Percent_Complete * i.Weight), 0) / totalWeight)
    : items.length > 0
      ? Math.round(items.reduce((s, i) => s + i.Percent_Complete, 0) / items.length)
      : 0;

  // ─── Command Bar ─────────────────────────────────────────────────────────

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'importWbs',
      text: isImporting ? 'Importing...' : 'Import WBS',
      iconProps: { iconName: 'ExcelDocument' },
      disabled: !projectCode || isImporting,
      onClick: handleImportButtonClick,
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        fetchScheduleItems().catch((error) => console.error('Error refreshing schedule items:', error));
        onRefresh();
      },
    },
  ];

  // Hidden file input
  const hiddenFileInput = (
    <input
      ref={fileInputRef}
      type="file"
      accept=".xlsx,.xls,.csv"
      style={{ display: 'none' }}
      onChange={handleFileSelected}
    />
  );

  // ─── Columns ─────────────────────────────────────────────────────────────

  const columns: IColumn[] = [
    { key: 'wbs', name: 'WBS ID', fieldName: 'WBS_ID', minWidth: 60, maxWidth: 80 },
    {
      key: 'activity', name: 'Activity', fieldName: 'Activity_Name', minWidth: 180, maxWidth: 300,
      onRender: (item: IScheduleItem) => (
        <span style={{ paddingLeft: `${(item.Level - 1) * 16}px`, fontWeight: item.Level === 1 ? 600 : 400 }}>
          {item.Activity_Name}
        </span>
      ),
    },
    { key: 'duration', name: 'Duration', fieldName: 'Duration_Days', minWidth: 60, maxWidth: 80, onRender: (item: IScheduleItem) => <span>{item.Duration_Days ? `${item.Duration_Days}d` : '—'}</span> },
    { key: 'pStart', name: 'Plan Start', minWidth: 90, maxWidth: 110, onRender: (item: IScheduleItem) => <span>{formatDate(item.Planned_Start)}</span> },
    { key: 'pFinish', name: 'Plan Finish', minWidth: 90, maxWidth: 110, onRender: (item: IScheduleItem) => <span>{formatDate(item.Planned_Finish)}</span> },
    { key: 'aStart', name: 'Act. Start', minWidth: 90, maxWidth: 110, onRender: (item: IScheduleItem) => <span>{formatDate(item.Actual_Start)}</span> },
    { key: 'aFinish', name: 'Act. Finish', minWidth: 90, maxWidth: 110, onRender: (item: IScheduleItem) => <span>{formatDate(item.Actual_Finish)}</span> },
    {
      key: 'progress', name: '% Complete', minWidth: 120, maxWidth: 150,
      onRender: (item: IScheduleItem) => (
        <div className={styles.progressCell}>
          <ProgressIndicator
            percentComplete={item.Percent_Complete / 100}
            barHeight={6}
            styles={{ root: { flex: 1 }, progressBar: { backgroundColor: item.Status === 'Delayed' ? '#EF4444' : '#2563EB' } }}
          />
          <span style={{ fontSize: '12px', fontWeight: 600, minWidth: '32px' }}>{item.Percent_Complete}%</span>
        </div>
      ),
    },
    {
      key: 'status', name: 'Status', minWidth: 90, maxWidth: 110,
      onRender: (item: IScheduleItem) => {
        const statusStyle = getStatusStyle(item.Status);
        return (
          <span className={styles.statusPill} style={{ backgroundColor: statusStyle.bg, color: statusStyle.color }}>
            {item.Status}
          </span>
        );
      },
    },
    {
      key: 'variance', name: 'Variance', minWidth: 70, maxWidth: 90,
      onRender: (item: IScheduleItem) => {
        if (item.Status === 'Completed' || item.Status === 'Not Started') return <span>—</span>;
        const days = getDaysVariance(item.Planned_Finish, item.Actual_Finish);
        if (days === null) return <span>—</span>;
        if (days <= 0) return <span style={{ color: '#10B981', fontWeight: 500 }}>On time</span>;
        return <span className={styles.delayBadge}>+{days}d late</span>;
      },
    },
  ];

  // ─── Render ──────────────────────────────────────────────────────────────

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Loading schedule..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (items.length === 0 && !importResult && !isImporting) {
    return (
      <Stack horizontalAlign="center" tokens={{ childrenGap: 12, padding: 40 }}>
        <Icon iconName="Timeline" styles={{ root: { fontSize: 48, color: '#D1D5DB' } }} />
        <Text variant="large" styles={{ root: { color: '#6B7280' } }}>No schedule data uploaded</Text>
        {storedDocument ? (
          <>
            <Text variant="small" styles={{ root: { color: '#9CA3AF' } }}>
              Source file found at {getProjectDocumentDisplayPath(PROJECT_WBS_FOLDER_NAME, projectCode, storedDocument.name)}.
            </Text>
            <PrimaryButton text="Import WBS File" iconProps={{ iconName: 'ExcelDocument' }} onClick={handleImportButtonClick} />
            {hiddenFileInput}
          </>
        ) : (
          <>
            <Text variant="small" styles={{ root: { color: '#9CA3AF' } }}>
              Upload a WBS file named {getProjectDocumentFileName(projectCode, 'schedule')} to {getProjectDocumentDisplayPath(PROJECT_WBS_FOLDER_NAME, projectCode, getProjectDocumentFileName(projectCode, 'schedule'))}.
            </Text>
            <PrimaryButton text="Upload & Import WBS" iconProps={{ iconName: 'ExcelDocument' }} onClick={handleImportButtonClick} />
            {hiddenFileInput}
          </>
        )}
      </Stack>
    );
  }

  return (
    <div className={styles.container}>
      {hiddenFileInput}

      <CommandBar items={commandBarItems} />

      {/* Import Progress */}
      {isImporting && importProgress && (
        <div style={{ padding: '12px 16px', backgroundColor: '#F0F6FF', borderRadius: '8px', border: '1px solid #B3D4FF' }}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
            <Spinner size={SpinnerSize.small} />
            <Text variant="small" styles={{ root: { color: '#1E2532', fontWeight: 500 } }}>
              {importProgress.message}
            </Text>
          </Stack>
          {importProgress.totalRows && importProgress.totalRows > 0 && (
            <ProgressIndicator percentComplete={importProgress.progress / 100} barHeight={4} styles={{ root: { marginTop: 8 } }} />
          )}
        </div>
      )}

      {/* Import Result */}
      {importResult && (
        <MessageBar messageBarType={importResult.success ? MessageBarType.success : MessageBarType.error} isMultiline onDismiss={() => setImportResult(null)}>
          <strong>{importResult.summary}</strong>
          {importResult.details.length > 0 && (
            <ul style={{ margin: '8px 0 0', paddingLeft: 20 }}>
              {importResult.details.slice(0, 10).map((d, i) => (
                <li key={i}>
                  <strong>{d.wbsId}</strong>: {d.action === 'created' ? '✅ Created' : d.action === 'updated' ? '🔄 Updated' : '⏭️ Skipped'}
                  {d.reason ? ` — ${d.reason}` : ''}
                </li>
              ))}
              {importResult.details.length > 10 && (
                <li style={{ color: '#6B7280' }}>...and {importResult.details.length - 10} more rows</li>
              )}
            </ul>
          )}
        </MessageBar>
      )}

      {/* Import Error */}
      {importError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setImportError(null)}>
          {importError}
        </MessageBar>
      )}

      {/* Summary Cards */}
      <div className={styles.summaryRow}>
        <div className={styles.summaryCard}>
          <Icon iconName="TaskList" styles={{ root: { fontSize: 24, color: '#2563EB' } }} />
          <div>
            <div className={styles.cardValue}>{totalActivities}</div>
            <div className={styles.cardLabel}>Total Activities</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="CompletedSolid" styles={{ root: { fontSize: 24, color: '#10B981' } }} />
          <div>
            <div className={styles.cardValue}>{completed}</div>
            <div className={styles.cardLabel}>Completed</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="Running" styles={{ root: { fontSize: 24, color: '#2563EB' } }} />
          <div>
            <div className={styles.cardValue}>{inProgress}</div>
            <div className={styles.cardLabel}>In Progress</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="Warning" styles={{ root: { fontSize: 24, color: '#EF4444' } }} />
          <div>
            <div className={styles.cardValue} style={{ color: delayed > 0 ? '#EF4444' : '#1E2532' }}>{delayed}</div>
            <div className={styles.cardLabel}>Delayed</div>
          </div>
        </div>
      </div>

      {/* Overall Progress */}
      <div className={styles.overallProgress}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
          <Text styles={{ root: { fontWeight: 600, color: '#1E2532' } }}>Overall Progress</Text>
          <div style={{ flex: 1 }}>
            <ProgressIndicator
              percentComplete={weightedComplete / 100}
              barHeight={10}
              styles={{ progressBar: { backgroundColor: '#2563EB' } }}
            />
          </div>
          <Text styles={{ root: { fontWeight: 700, fontSize: 18, color: '#2563EB' } }}>{weightedComplete}%</Text>
        </Stack>
      </div>

      {/* Activities List */}
      <Text className={styles.sectionTitle}>Activities</Text>
      <DetailsList
        items={items}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        compact={isMobileView}
      />
    </div>
  );
};

export default ProjectScheduleTab;
