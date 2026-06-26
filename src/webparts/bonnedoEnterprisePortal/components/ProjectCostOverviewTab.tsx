import * as React from 'react';
import {
  Text, getTheme, mergeStyleSets, IconButton, Icon, Spinner, SpinnerSize,
  Dropdown, IDropdownOption,
  DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, ConstrainMode,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface IProjectCostOverviewTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IPhaseCostSummary {
  phase: string;
  budgetAmount: number;
  actualAmount: number;
  commitmentAmount: number;
  variance: number;
  utilization: number;
}

const ProjectCostOverviewTab: React.FC<IProjectCostOverviewTabProps> = ({
  spHttpClient,
  pageContext,
  isMobileView,
  onRefresh,
}) => {
  const theme = getTheme();
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [projects, setProjects] = React.useState<IDropdownOption[]>([]);
  const [selectedProject, setSelectedProject] = React.useState<string>('');
  const [phaseData, setPhaseData] = React.useState<IPhaseCostSummary[]>([]);
  const [totals, setTotals] = React.useState({ budget: 0, actual: 0, commitment: 0, variance: 0, utilization: 0 });
  const [refreshKey, setRefreshKey] = React.useState(0);

  const classNames = mergeStyleSets({
    container: { padding: '20px', height: '100%', display: 'flex', flexDirection: 'column' },
    header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' as const },
    headerTitle: { flex: 1, minWidth: '200px' },
    headerActions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
    summaryRow: { display: 'flex', gap: '16px', flexWrap: 'wrap', marginBottom: '16px' },
    summaryCard: {
      flex: 1, minWidth: '150px', padding: '16px', backgroundColor: theme.palette.white,
      borderRadius: '8px', border: `1px solid ${theme.palette.neutralLight}`,
      boxShadow: theme.effects.elevation4,
    },
    summaryLabel: { fontSize: '12px', color: theme.palette.neutralSecondary, marginBottom: '4px' },
    summaryValue: { fontSize: '22px', fontWeight: 700, color: theme.palette.neutralPrimary },
    gridContainer: { flex: 1, overflow: 'auto' as const, minHeight: 0, backgroundColor: theme.palette.white, border: `1px solid ${theme.palette.neutralLight}`, borderRadius: '4px' },
    loadingContainer: { display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px', color: theme.palette.neutralSecondary },
    emptyState: { textAlign: 'center', padding: '60px 20px', color: theme.palette.neutralSecondary },
  });

  const loadProjects = React.useCallback(async (): Promise<void> => {
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$select=Project_Code,Project_Name&$top=200&$orderby=Project_Name`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const opts: IDropdownOption[] = (data.value || []).map((p: any) => ({
          key: p.Project_Code || '',
          text: `${p.Project_Code || ''} \u2014 ${p.Project_Name || ''}`,
        })).filter((o: IDropdownOption) => o.key);
        setProjects(opts);
      }
    } catch { /* silent */ }
  }, [spHttpClient, pageContext]);

  const loadPhaseData = React.useCallback(async (): Promise<void> => {
    if (!selectedProject) { setLoading(false); return; }
    setLoading(true);
    setError(null);
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const escaped = selectedProject.replace(/'/g, "''");
      let budgetMap = new Map<string, number>();
      let costMap = new Map<string, number>();
      let poCommitMap = new Map<string, number>();

      // Budget items by phase
      try {
        const bUrl = `${webUrl}/_api/web/lists/getByTitle('Project_Budget_Items')/items?$filter=Project_Code eq '${escaped}'&$select=CBS_Code,Phase,Budget_Amount&$top=500`;
        const bResp = await spHttpClient.get(bUrl, SPHttpClient.configurations.v1);
        if (bResp.ok) {
          const bData = await bResp.json();
          (bData.value || []).forEach((item: any) => {
            const phase = item.Phase || (item.CBS_Code ? item.CBS_Code.split('-')[0] : 'General') || 'General';
            budgetMap.set(phase, (budgetMap.get(phase) || 0) + (item.Budget_Amount || 0));
          });
        }
      } catch { /* silent */ }

      // Cost link actuals
      try {
        const cUrl = `${webUrl}/_api/web/lists/getByTitle('Cost_Transactions')/items?$filter=Project_Code eq '${escaped}'&$select=Amount,Transaction_Type,Phase&$top=500`;
        const cResp = await spHttpClient.get(cUrl, SPHttpClient.configurations.v1);
        if (cResp.ok) {
          const cData = await cResp.json();
          (cData.value || []).forEach((item: any) => {
            const phase = item.Phase || 'General';
            const amt = item.Amount || 0;
            const txnType = (item.Transaction_Type || '').toLowerCase();
            if (txnType.includes('po commitment')) {
              poCommitMap.set(phase, (poCommitMap.get(phase) || 0) + amt);
            } else {
              costMap.set(phase, (costMap.get(phase) || 0) + amt);
            }
          });
        }
      } catch { /* silent */ }

      // Merge phases
      const phaseSet: string[] = [];
      const addPhase = (p: string): void => { if (phaseSet.indexOf(p) < 0) phaseSet.push(p); };
      budgetMap.forEach((_, k) => addPhase(k));
      costMap.forEach((_, k) => addPhase(k));
      poCommitMap.forEach((_, k) => addPhase(k));
      if (phaseSet.length === 0) phaseSet.push('General');

      const summaries: IPhaseCostSummary[] = [];
      let totBudget = 0, totActual = 0, totCommit = 0;
      for (let pi = 0; pi < phaseSet.length; pi++) {
        const phase = phaseSet[pi];
        const budget = budgetMap.get(phase) || 0;
        const actual = costMap.get(phase) || 0;
        const commit = poCommitMap.get(phase) || 0;
        const variance = budget > 0 ? budget - actual - commit : 0;
        const utilization = budget > 0 ? Math.round(((actual + commit) / budget) * 100) : 0;
        summaries.push({ phase, budgetAmount: budget, actualAmount: actual, commitmentAmount: commit, variance, utilization });
        totBudget += budget; totActual += actual; totCommit += commit;
      }
      summaries.sort((a, b) => a.phase.localeCompare(b.phase));

      setPhaseData(summaries);
      setTotals({
        budget: totBudget, actual: totActual, commitment: totCommit,
        variance: totBudget > 0 ? totBudget - totActual - totCommit : 0,
        utilization: totBudget > 0 ? Math.round(((totActual + totCommit) / totBudget) * 100) : 0,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load cost data');
    } finally {
      setLoading(false);
    }
  }, [spHttpClient, pageContext, selectedProject]);

  React.useEffect(() => { loadProjects().catch(() => undefined); }, [loadProjects]);
  React.useEffect(() => { loadPhaseData().catch(() => undefined); }, [loadPhaseData, refreshKey]);

  const handleRefresh = (): void => { setRefreshKey(prev => prev + 1); onRefresh(); };

  const fmt = (n: number): string => {
    if (n >= 1000000) return `\u20A6${(n / 1000000).toFixed(1)}M`;
    if (n >= 1000) return `\u20A6${(n / 1000).toFixed(1)}K`;
    return `\u20A6${n.toFixed(0)}`;
  };

  const getBarColor = (utilization: number): string => {
    if (utilization > 100) return '#D13438';
    if (utilization > 80) return '#F39C12';
    return '#0078D4';
  };

  const phaseColumns: IColumn[] = [
    {
      key: 'phase', name: 'Phase', fieldName: 'phase', minWidth: 120, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => item ? <span style={{ fontWeight: 600 }}>{item.phase}</span> : null,
    },
    {
      key: 'budgetAmount', name: 'Budget', fieldName: 'budgetAmount', minWidth: 120, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => item ? <span>{fmt(item.budgetAmount)}</span> : null,
    },
    {
      key: 'actualAmount', name: 'Actual', fieldName: 'actualAmount', minWidth: 120, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => item ? <span style={{ color: '#D13438' }}>{fmt(item.actualAmount)}</span> : null,
    },
    {
      key: 'commitmentAmount', name: 'Commitment', fieldName: 'commitmentAmount', minWidth: 120, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => item ? <span style={{ color: '#F39C12' }}>{fmt(item.commitmentAmount)}</span> : null,
    },
    {
      key: 'variance', name: 'Variance', fieldName: 'variance', minWidth: 120, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => item ? (
        <span style={{
          fontWeight: 600, color: item.variance < 0 ? '#A80000' : '#107C10',
          backgroundColor: item.variance < 0 ? '#FDE7E9' : '#DFF6DD',
          padding: '2px 8px', borderRadius: '10px', fontSize: '11px',
        }}>
          {item.variance < 0 ? `-${fmt(Math.abs(item.variance))}` : fmt(item.variance)}
        </span>
      ) : null,
    },
    {
      key: 'utilization', name: 'Utilization', fieldName: 'utilization', minWidth: 140, isResizable: true,
      onRender: (item?: IPhaseCostSummary) => {
        if (!item) return null;
        const pct = Math.min(item.utilization, 100);
        const color = getBarColor(item.utilization);
        return (
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, width: '100%' }}>
            <span style={{ flex: 1, height: 8, backgroundColor: '#EDEDED', borderRadius: 4, overflow: 'hidden', display: 'inline-block', maxWidth: 100 }}>
              <span style={{ display: 'block', height: 8, borderRadius: 4, width: `${pct}%`, backgroundColor: color, transition: 'width 0.3s ease' }} />
            </span>
            <span style={{ fontSize: 12, fontWeight: 600, color }}>{item.utilization}%</span>
          </span>
        );
      },
    },
  ];

  return (
    <div className={classNames.container}>
      <div className={classNames.header}>
        <div className={classNames.headerTitle}>
          <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
            Cost Overview
          </Text>
          <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
            Budget vs actual spending by phase
          </Text>
        </div>
        <div className={classNames.headerActions}>
          <Dropdown
            placeholder="Select project..."
            options={projects}
            selectedKey={selectedProject}
            onChange={(_, opt) => setSelectedProject(opt?.key as string || '')}
            styles={{ root: { minWidth: '260px' } }}
          />
          <IconButton iconProps={{ iconName: 'Refresh' }} onClick={handleRefresh} title="Refresh" />
        </div>
      </div>

      {selectedProject && !loading && phaseData.length > 0 && (
        <div className={classNames.summaryRow}>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Total Budget</div>
            <div className={classNames.summaryValue}>{fmt(totals.budget)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Actual Spent</div>
            <div className={classNames.summaryValue} style={{ color: '#D13438' }}>{fmt(totals.actual)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Committed (PO)</div>
            <div className={classNames.summaryValue} style={{ color: '#F39C12' }}>{fmt(totals.commitment)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Variance</div>
            <div className={classNames.summaryValue} style={{ color: totals.variance < 0 ? '#D13438' : '#107C10' }}>{fmt(totals.variance)}</div>
          </div>
          <div className={classNames.summaryCard}>
            <div className={classNames.summaryLabel}>Utilization</div>
            <div className={classNames.summaryValue} style={{ color: totals.utilization > 100 ? '#D13438' : totals.utilization > 80 ? '#F39C12' : '#107C10' }}>
              {totals.utilization}%
            </div>
          </div>
        </div>
      )}

      {loading ? (
        <div className={classNames.loadingContainer}>
          <Spinner size={SpinnerSize.medium} label="Loading cost data..." />
        </div>
      ) : !selectedProject ? (
        <div className={classNames.emptyState}>
          <Icon iconName="StackedBarChart" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">Select a project</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>Choose a project to view cost overview by phase.</Text>
        </div>
      ) : phaseData.length === 0 ? (
        <div className={classNames.emptyState}>
          <Icon iconName="StackedBarChart" style={{ fontSize: 48, color: '#CCC', display: 'block', marginBottom: 12 }} />
          <Text variant="large">No cost data found</Text>
          <Text variant="small" style={{ display: 'block', marginTop: 4 }}>No budget or cost records for the selected project.</Text>
        </div>
      ) : (
        <div className={classNames.gridContainer}>
          <DetailsList
            items={phaseData}
            columns={phaseColumns}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            constrainMode={ConstrainMode.horizontalConstrained}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
            compact={true}
          />
        </div>
      )}
    </div>
  );
};

export default ProjectCostOverviewTab;
