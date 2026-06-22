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
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
  getProjectDocumentDisplayPath,
  getProjectDocumentFileName,
  getProjectDocumentLibraryRootUrl,
  getStoredProjectDocument,
  IStoredProjectDocument,
  PROJECT_BUDGET_FOLDER_NAME,
} from '../../services/ProjectDocumentStorageService';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IProjectBudgetTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  projectCode: string;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IBudgetItem {
  ID: number;
  Project_Code: string;
  CBS_Code: string;
  Phase: string;
  Description: string;
  UOM: string;
  Qty: number;
  Rate: number;
  Budget_Amount: number;
  Original_Budget: number;
  Actual_Amount: number;
  Variance_Budget: number;
  Remarks: string;
}

interface IPhaseSummary {
  phase: string;
  budgetAmount: number;
  originalBudget: number;
  actualAmount: number;
  variance: number;
  utilizationPct: number;
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const getStyles = () =>
  mergeStyleSets({
    container: {
      display: 'flex',
      flexDirection: 'column' as const,
      gap: '16px',
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
      minWidth: '180px',
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
    phaseTable: {
      backgroundColor: '#FFFFFF',
      borderRadius: '8px',
      border: '1px solid #E5E7EB',
      overflow: 'hidden',
    },
    phaseRow: {
      display: 'flex',
      padding: '12px 16px',
      borderBottom: '1px solid #F3F4F6',
      alignItems: 'center',
    },
    phaseHeader: {
      display: 'flex',
      padding: '12px 16px',
      backgroundColor: '#1E2532',
      color: '#FFFFFF',
      fontWeight: '600',
      fontSize: '12px',
    },
    statusPill: {
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
      display: 'inline-block',
    },
    sectionTitle: {
      fontSize: '16px',
      fontWeight: '600',
      color: '#1E2532',
      marginTop: '8px',
    },
  });

// ─── Helper Functions ────────────────────────────────────────────────────────

const formatCurrency = (value: number): string => {
  if (value >= 1000000000) return `₦${(value / 1000000000).toFixed(2)}B`;
  if (value >= 1000000) return `₦${(value / 1000000).toFixed(2)}M`;
  if (value >= 1000) return `₦${(value / 1000).toFixed(1)}K`;
  return `₦${value.toLocaleString()}`;
};

const getVarianceStatus = (budget: number, actual: number): { label: string; color: string; bg: string } => {
  if (actual === 0 && budget > 0) return { label: 'No Spend', color: '#6B7280', bg: '#F3F4F6' };
  const pct = budget > 0 ? (actual / budget) * 100 : 0;
  if (pct > 100) return { label: 'Over Budget', color: '#FFFFFF', bg: '#EF4444' };
  if (pct >= 90) return { label: 'At Risk', color: '#FFFFFF', bg: '#F59E0B' };
  return { label: 'On Track', color: '#FFFFFF', bg: '#10B981' };
};

// ─── Component ───────────────────────────────────────────────────────────────

const ProjectBudgetTab: React.FC<IProjectBudgetTabProps> = ({
  spHttpClient,
  pageContext,
  projectCode,
  isMobileView,
  onRefresh,
}) => {
  const styles = getStyles();
  const [items, setItems] = React.useState<IBudgetItem[]>([]);
  const [storedDocument, setStoredDocument] = React.useState<IStoredProjectDocument | undefined>(undefined);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  // ─── Fetch Data ──────────────────────────────────────────────────────────

  const fetchBudgetItems = React.useCallback(async (): Promise<void> => {
    if (!projectCode) {
      setItems([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const escapedProjectCode = projectCode.replace(/'/g, "''");
      const filteredUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Budget_Items')/items?$filter=Project_Code eq '${escapedProjectCode}'&$orderby=CBS_Code asc&$top=500`;
      const fallbackUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Budget_Items')/items?$top=5000`;
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

      if (!response.ok) throw new Error(`Failed to fetch budget items: ${response.status}`);

      const data = await response.json();
      const fetched: IBudgetItem[] = (data.value || [])
        .filter((item: any) => String(item.Project_Code || '').trim() === projectCode.trim())
        .map((item: any) => ({
        ID: item.ID,
        Project_Code: item.Project_Code || '',
        CBS_Code: item.CBS_Code || '',
        Phase: item.Phase || '',
        Description: item.Description || item.Title || '',
        UOM: item.UOM || '',
        Qty: item.Qty || 0,
        Rate: item.Rate || 0,
        Budget_Amount: item.Budget_Amount || 0,
        Original_Budget: item.Original_Budget || 0,
        Actual_Amount: item.Actual_Amount || 0,
        Variance_Budget: item.Variance_Budget || 0,
        Remarks: item.Remarks || '',
      }));

      setItems(fetched);
    } catch (err) {
      console.error('Error fetching budget items:', err);
      setError(err instanceof Error ? err.message : 'Failed to load budget data');
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
        PROJECT_BUDGET_FOLDER_NAME,
        projectCode,
        'budget'
      );

      setStoredDocument(document);
    } catch (error) {
      console.warn('[ProjectBudgetTab] Unable to read stored CBS source file:', error);
      setStoredDocument(undefined);
    }
  }, [spHttpClient, pageContext, projectCode]);

  React.useEffect(() => {
    fetchBudgetItems().catch((error) => console.error('Error refreshing budget items:', error));
    fetchStoredDocument().catch((error) => console.error('Error refreshing stored budget document:', error));
  }, [fetchBudgetItems, fetchStoredDocument]);

  // ─── Calculations ────────────────────────────────────────────────────────

  const totalBudget = items.reduce((sum, i) => sum + i.Budget_Amount, 0);
  const totalOriginal = items.reduce((sum, i) => sum + i.Original_Budget, 0);
  const totalActual = items.reduce((sum, i) => sum + i.Actual_Amount, 0);
  const totalVariance = totalBudget - totalActual;
  const utilizationPct = totalBudget > 0 ? Math.round((totalActual / totalBudget) * 100) : 0;

  // Phase summary
  const phases = ['MOBILIZATION', 'ENGINEERING', 'CONSTRUCTION', 'REINSTATEMENT', 'EQUIPMENT', 'LABOUR'];
  const phaseSummaries: IPhaseSummary[] = phases.map((phase) => {
    const phaseItems = items.filter((i) => i.Phase === phase);
    const budget = phaseItems.reduce((s, i) => s + i.Budget_Amount, 0);
    const original = phaseItems.reduce((s, i) => s + i.Original_Budget, 0);
    const actual = phaseItems.reduce((s, i) => s + i.Actual_Amount, 0);
    return {
      phase,
      budgetAmount: budget,
      originalBudget: original,
      actualAmount: actual,
      variance: budget - actual,
      utilizationPct: budget > 0 ? Math.round((actual / budget) * 100) : 0,
    };
  });

  // ─── Command Bar ─────────────────────────────────────────────────────────

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        fetchBudgetItems().catch((error) => console.error('Error refreshing budget items:', error));
        onRefresh();
      },
    },
  ];

  // ─── Columns ─────────────────────────────────────────────────────────────

  const columns: IColumn[] = [
    { key: 'cbs', name: 'CBS Code', fieldName: 'CBS_Code', minWidth: 70, maxWidth: 90 },
    { key: 'phase', name: 'Phase', fieldName: 'Phase', minWidth: 100, maxWidth: 130 },
    { key: 'desc', name: 'Description', fieldName: 'Description', minWidth: 180, maxWidth: 300, isMultiline: true },
    { key: 'uom', name: 'UOM', fieldName: 'UOM', minWidth: 50, maxWidth: 60 },
    { key: 'qty', name: 'Qty', fieldName: 'Qty', minWidth: 50, maxWidth: 70, onRender: (item: IBudgetItem) => <span>{item.Qty.toLocaleString()}</span> },
    { key: 'rate', name: 'Rate', fieldName: 'Rate', minWidth: 80, maxWidth: 100, onRender: (item: IBudgetItem) => <span>{formatCurrency(item.Rate)}</span> },
    { key: 'budget', name: 'Budget', fieldName: 'Budget_Amount', minWidth: 90, maxWidth: 120, onRender: (item: IBudgetItem) => <span style={{ fontWeight: 600 }}>{formatCurrency(item.Budget_Amount)}</span> },
    { key: 'original', name: 'Original', fieldName: 'Original_Budget', minWidth: 90, maxWidth: 120, onRender: (item: IBudgetItem) => <span>{formatCurrency(item.Original_Budget)}</span> },
    { key: 'actual', name: 'Actual', fieldName: 'Actual_Amount', minWidth: 90, maxWidth: 120, onRender: (item: IBudgetItem) => <span>{formatCurrency(item.Actual_Amount)}</span> },
    {
      key: 'status', name: 'Status', minWidth: 90, maxWidth: 110,
      onRender: (item: IBudgetItem) => {
        const status = getVarianceStatus(item.Budget_Amount, item.Actual_Amount);
        return (
          <span className={styles.statusPill} style={{ backgroundColor: status.bg, color: status.color }}>
            {status.label}
          </span>
        );
      },
    },
  ];

  // ─── Render ──────────────────────────────────────────────────────────────

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Loading budget data..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (items.length === 0) {
    return (
      <Stack horizontalAlign="center" tokens={{ childrenGap: 12, padding: 40 }}>
        <Icon iconName="Money" styles={{ root: { fontSize: 48, color: '#D1D5DB' } }} />
        <Text variant="large" styles={{ root: { color: '#6B7280' } }}>No budget data uploaded</Text>
        {storedDocument ? (
          <>
            <Text variant="small" styles={{ root: { color: '#9CA3AF' } }}>
              Source file found at {getProjectDocumentDisplayPath(PROJECT_BUDGET_FOLDER_NAME, projectCode, storedDocument.name)}.
            </Text>
            <Text variant="small" styles={{ root: { color: '#9CA3AF' } }}>
              Parsed list data is not available yet. Refresh after the CBS processing flow completes.
            </Text>
          </>
        ) : (
          <Text variant="small" styles={{ root: { color: '#9CA3AF' } }}>
            Upload a CBS file named {getProjectDocumentFileName(projectCode, 'budget')} to {getProjectDocumentDisplayPath(PROJECT_BUDGET_FOLDER_NAME, projectCode, getProjectDocumentFileName(projectCode, 'budget'))}.
          </Text>
        )}
      </Stack>
    );
  }

  return (
    <div className={styles.container}>
      <CommandBar items={commandBarItems} />

      {/* Summary Cards */}
      <div className={styles.summaryRow}>
        <div className={styles.summaryCard}>
          <Icon iconName="Money" styles={{ root: { fontSize: 24, color: '#2563EB' } }} />
          <div>
            <div className={styles.cardValue}>{formatCurrency(totalBudget)}</div>
            <div className={styles.cardLabel}>Total Budget (Revised)</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="DocumentSet" styles={{ root: { fontSize: 24, color: '#6B7280' } }} />
          <div>
            <div className={styles.cardValue}>{formatCurrency(totalOriginal)}</div>
            <div className={styles.cardLabel}>Original Budget</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="PaymentCard" styles={{ root: { fontSize: 24, color: '#F59E0B' } }} />
          <div>
            <div className={styles.cardValue}>{formatCurrency(totalActual)}</div>
            <div className={styles.cardLabel}>Actual Spend</div>
          </div>
        </div>
        <div className={styles.summaryCard}>
          <Icon iconName="AreaChart" styles={{ root: { fontSize: 24, color: totalVariance >= 0 ? '#10B981' : '#EF4444' } }} />
          <div>
            <div className={styles.cardValue} style={{ color: totalVariance >= 0 ? '#10B981' : '#EF4444' }}>
              {utilizationPct}%
            </div>
            <div className={styles.cardLabel}>Budget Utilization</div>
          </div>
        </div>
      </div>

      {/* Phase Breakdown */}
      <Text className={styles.sectionTitle}>Phase Breakdown</Text>
      <div className={styles.phaseTable}>
        <div className={styles.phaseHeader}>
          <span style={{ flex: 2 }}>Phase</span>
          <span style={{ flex: 1, textAlign: 'right' }}>Budget</span>
          <span style={{ flex: 1, textAlign: 'right' }}>Actual</span>
          <span style={{ flex: 1, textAlign: 'right' }}>Variance</span>
          <span style={{ flex: 1, textAlign: 'right' }}>Utilization</span>
          <span style={{ flex: 1, textAlign: 'center' }}>Status</span>
        </div>
        {phaseSummaries.filter(p => p.budgetAmount > 0).map((ps) => {
          const status = getVarianceStatus(ps.budgetAmount, ps.actualAmount);
          return (
            <div key={ps.phase} className={styles.phaseRow}>
              <span style={{ flex: 2, fontWeight: 500 }}>{ps.phase}</span>
              <span style={{ flex: 1, textAlign: 'right' }}>{formatCurrency(ps.budgetAmount)}</span>
              <span style={{ flex: 1, textAlign: 'right' }}>{formatCurrency(ps.actualAmount)}</span>
              <span style={{ flex: 1, textAlign: 'right', color: ps.variance >= 0 ? '#10B981' : '#EF4444' }}>
                {formatCurrency(Math.abs(ps.variance))}
              </span>
              <span style={{ flex: 1, textAlign: 'right' }}>{ps.utilizationPct}%</span>
              <span style={{ flex: 1, textAlign: 'center' }}>
                <span className={styles.statusPill} style={{ backgroundColor: status.bg, color: status.color }}>
                  {status.label}
                </span>
              </span>
            </div>
          );
        })}
      </div>

      {/* Line Items */}
      <Text className={styles.sectionTitle}>Budget Line Items</Text>
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

export default ProjectBudgetTab;
