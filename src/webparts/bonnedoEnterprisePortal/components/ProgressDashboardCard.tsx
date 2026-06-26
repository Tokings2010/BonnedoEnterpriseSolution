import * as React from 'react';
import {
  Text,
  getTheme,
  mergeStyleSets,
  Icon,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface IProgressDashboardCardProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  projectCode: string;
}

interface IProgressData {
  physicalProgress: number;
  costProgress: number;
  delayedCount: number;
  budgetStatus: { budget: number; spent: number; remaining: number };
  earnedValue: { plannedValue: number; earnedValue: number; actualCost: number; spi: number; cpi: number };
  isLoading: boolean;
  error: string | null;
}

const ProgressDashboardCard: React.FC<IProgressDashboardCardProps> = ({
  spHttpClient,
  pageContext,
  projectCode,
}) => {
  const theme = getTheme();
  const [data, setData] = React.useState<IProgressData>({
    physicalProgress: 0,
    costProgress: 0,
    delayedCount: 0,
    budgetStatus: { budget: 0, spent: 0, remaining: 0 },
    earnedValue: { plannedValue: 0, earnedValue: 0, actualCost: 0, spi: 1, cpi: 1 },
    isLoading: true,
    error: null,
  });

  const classNames = mergeStyleSets({
    card: {
      padding: '20px',
      backgroundColor: theme.palette.white,
      borderRadius: '12px',
      boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
      border: `1px solid ${theme.palette.neutralLighter}`,
    },
    cardHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '16px',
    },
    cardTitle: {
      fontSize: '15px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
    },
    grid: {
      display: 'grid',
      gridTemplateColumns: '1fr 1fr',
      gap: '16px',
    },
    metricBox: {
      display: 'flex',
      flexDirection: 'column',
      gap: '4px',
    },
    metricLabel: {
      fontSize: '11px',
      fontWeight: 600,
      color: theme.palette.neutralSecondary,
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
    },
    metricValue: {
      fontSize: '22px',
      fontWeight: 700,
      color: theme.palette.neutralPrimary,
    },
    progressBarOuter: {
      width: '100%',
      height: '8px',
      backgroundColor: theme.palette.neutralLight,
      borderRadius: '4px',
      overflow: 'hidden',
      marginTop: '4px',
    },
    progressBarInner: {
      height: '100%',
      borderRadius: '4px',
      transition: 'width 0.5s ease',
    },
    divider: {
      height: '1px',
      backgroundColor: theme.palette.neutralLight,
      margin: '12px 0',
    },
    row: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '8px',
    },
    label: {
      fontSize: '12px',
      color: theme.palette.neutralSecondary,
    },
    value: {
      fontSize: '13px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
    },
    evRow: {
      display: 'flex',
      gap: '12px',
      flexWrap: 'wrap',
      marginTop: '8px',
    },
    evItem: {
      flex: 1,
      minWidth: '80px',
      padding: '8px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '6px',
      textAlign: 'center',
    },
    evLabel: {
      fontSize: '10px',
      fontWeight: 600,
      color: theme.palette.neutralSecondary,
      textTransform: 'uppercase',
    },
    evValue: {
      fontSize: '16px',
      fontWeight: 700,
      marginTop: '2px',
    },
    statusIndicator: {
      display: 'flex',
      alignItems: 'center',
      gap: '6px',
      padding: '4px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: 600,
    },
    emptyState: {
      padding: '24px',
      textAlign: 'center',
      color: theme.palette.neutralSecondary,
    },
  });

  const fetchProgressData = React.useCallback(async (): Promise<void> => {
    if (!projectCode) {
      setData((prev) => ({ ...prev, isLoading: false }));
      return;
    }

    setData((prev) => ({ ...prev, isLoading: true, error: null }));
    try {
      const webUrl = pageContext.web.absoluteUrl;
      let physicalProgress = 0;
      let delayedCount = 0;
      let budgetAmount = 0;
      let costSpent = 0;

      // Fetch WBS activities from Project_Schedule (try with filter, fallback to client-side)
      try {
        const escapedProjectCode = projectCode.replace(/'/g, "''");
        const wbsUrl = `${webUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$filter=Project_Code eq '${escapedProjectCode}'&$select=Percent_Complete,Status&$top=500`;
        let wbsResp: SPHttpClientResponse = await spHttpClient.get(wbsUrl, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json;odata=nometadata' },
        });

        if (!wbsResp.ok) {
          // Fallback: fetch all and filter client-side
          const fallbackUrl = `${webUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$top=5000&$select=Percent_Complete,Status,Project_Code`;
          wbsResp = await spHttpClient.get(fallbackUrl, SPHttpClient.configurations.v1, {
            headers: { 'Accept': 'application/json;odata=nometadata' },
          });
          if (wbsResp.ok) {
            const allData = await wbsResp.json();
            const activities = (allData.value || []).filter(
              (a: any) => String(a.Project_Code || '').trim() === projectCode.trim()
            );
            if (activities.length > 0) {
              physicalProgress = Math.round(
                activities.reduce((s: number, a: any) => s + (a.Percent_Complete ?? 0), 0) / activities.length
              );
              delayedCount = activities.filter((a: any) => a.Status === 'Delayed').length;
            }
          }
        } else {
          const wbsData = await wbsResp.json();
          const activities = wbsData.value || [];
          if (activities.length > 0) {
            physicalProgress = Math.round(
              activities.reduce((s: number, a: any) => s + (a.Percent_Complete ?? 0), 0) / activities.length
            );
            delayedCount = activities.filter((a: any) => a.Status === 'Delayed').length;
          }
        }
      } catch {
        // WBS list may not exist yet - silently continue
      }

      // Fetch project budget from ENT_Project_Master (only use Contract_Value, not Budget_Amount)
      try {
        const escapedProjectCode = projectCode.replace(/'/g, "''");
        const budgetUrl = `${webUrl}/_api/web/lists/getbytitle('ENT_Project_Master')/items?$filter=Project_Code eq '${escapedProjectCode}'&$select=Contract_Value&$top=1`;
        const budgetResp: SPHttpClientResponse = await spHttpClient.get(budgetUrl, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json;odata=nometadata' },
        });
        if (budgetResp.ok) {
          const budgetData = await budgetResp.json();
          const proj = budgetData.value?.[0];
          budgetAmount = proj?.Contract_Value ?? 0;
        }
      } catch {
        // Project master list may not be accessible - silently continue
      }

      // Fetch cost records from Cost_Transactions (graceful if missing)
      try {
        const escapedProjectCode = projectCode.replace(/'/g, "''");
        const costUrl = `${webUrl}/_api/web/lists/getbytitle('Cost_Transactions')/items?$filter=Project_Code eq '${escapedProjectCode}'&$select=Amount&$top=500`;
        const costResp: SPHttpClientResponse = await spHttpClient.get(costUrl, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json;odata=nometadata' },
        });
        if (costResp.ok) {
          const costData = await costResp.json();
          costSpent = (costData.value || []).reduce((s: number, c: any) => s + (c.Amount || 0), 0);
        }
      } catch {
        // Cost Link list may not exist yet - silently continue
      }

      // Calculate cost progress
      const costProgress = budgetAmount > 0 ? Math.min(100, Math.round((costSpent / budgetAmount) * 100)) : 0;
      const remaining = Math.max(0, budgetAmount - costSpent);

      // Earned value estimates
      const plannedValue = budgetAmount > 0 ? Math.round(budgetAmount * (physicalProgress / 100)) : 0;
      const earnedValue = plannedValue;
      const actualCost = costSpent;
      const spi = physicalProgress > 0 ? physicalProgress / Math.max(1, physicalProgress) : 1;
      const cpi = earnedValue > 0 ? earnedValue / Math.max(1, actualCost) : 1;

      setData({
        physicalProgress,
        costProgress,
        delayedCount,
        budgetStatus: { budget: budgetAmount, spent: costSpent, remaining },
        earnedValue: { plannedValue, earnedValue, actualCost, spi, cpi },
        isLoading: false,
        error: null,
      });
    } catch (err) {
      console.error('[ProgressDashboardCard] Error:', err);
      setData((prev) => ({ ...prev, isLoading: false, error: 'Failed to load progress data' }));
    }
  }, [spHttpClient, pageContext, projectCode]);

  React.useEffect(() => {
    fetchProgressData().catch(() => undefined);
  }, [fetchProgressData]);

  const formatCurrency = (amount: number): string => {
    if (amount >= 1000000) return `\u20A6${(amount / 1000000).toFixed(1)}M`;
    if (amount >= 1000) return `\u20A6${(amount / 1000).toFixed(1)}K`;
    return `\u20A6${amount.toFixed(0)}`;
  };

  const getProgressColor = (pct: number): string => {
    if (pct >= 80) return '#107C10';
    if (pct >= 50) return '#0078D4';
    if (pct >= 25) return '#F39C12';
    return '#D13438';
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const getStatusText = (v: number, threshold: number, goodIfAbove: boolean): string => {
    if (goodIfAbove) return v >= threshold ? 'Good' : 'Needs Attention';
    return v <= threshold ? 'Good' : 'Needs Attention';
  };

  if (data.isLoading) {
    return (
      <div className={classNames.card}>
        <Spinner size={SpinnerSize.small} label="Loading progress data..." />
      </div>
    );
  }

  if (!projectCode) {
    return (
      <div className={classNames.emptyState}>
        <Icon iconName="TimelineProgress" style={{ fontSize: 28, marginBottom: 8, opacity: 0.4 }} />
        <Text variant="medium">Select a project to view progress</Text>
      </div>
    );
  }

  return (
    <div className={classNames.card}>
      <div className={classNames.cardHeader}>
        <div className={classNames.cardTitle}>
          <Icon iconName="TimelineProgress" style={{ marginRight: 6, color: theme.palette.themePrimary }} />
          Project Progress
        </div>
        {data.delayedCount > 0 && (
          <div className={classNames.statusIndicator} style={{ backgroundColor: '#FEE2E2', color: '#991B1B' }}>
            <Icon iconName="Warning" style={{ fontSize: 12 }} />
            {data.delayedCount} delayed
          </div>
        )}
      </div>

      <div className={classNames.grid}>
        {/* Physical Progress */}
        <div className={classNames.metricBox}>
          <span className={classNames.metricLabel}>Physical Progress</span>
          <div className={classNames.metricValue} style={{ color: getProgressColor(data.physicalProgress) }}>
            {data.physicalProgress}%
          </div>
          <div className={classNames.progressBarOuter}>
            <div
              className={classNames.progressBarInner}
              style={{ width: `${data.physicalProgress}%`, backgroundColor: getProgressColor(data.physicalProgress) }}
            />
          </div>
        </div>

        {/* Cost Progress */}
        <div className={classNames.metricBox}>
          <span className={classNames.metricLabel}>Cost Progress</span>
          <div className={classNames.metricValue} style={{ color: getProgressColor(data.costProgress) }}>
            {data.costProgress}%
          </div>
          <div className={classNames.progressBarOuter}>
            <div
              className={classNames.progressBarInner}
              style={{ width: `${data.costProgress}%`, backgroundColor: getProgressColor(data.costProgress) }}
            />
          </div>
        </div>
      </div>

      <div className={classNames.divider} />

      {/* Budget Status */}
      <div>
        <Text style={{ fontSize: '12px', fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: '8px', display: 'block' }}>
          Budget Status
        </Text>
        <div className={classNames.row}>
          <span className={classNames.label}>Budget</span>
          <span className={classNames.value}>{formatCurrency(data.budgetStatus.budget)}</span>
        </div>
        <div className={classNames.row}>
          <span className={classNames.label}>Spent</span>
          <span className={classNames.value} style={{ color: data.budgetStatus.spent > data.budgetStatus.budget ? '#D13438' : '#107C10' }}>
            {formatCurrency(data.budgetStatus.spent)}
          </span>
        </div>
        <div className={classNames.row}>
          <span className={classNames.label}>Remaining</span>
          <span className={classNames.value} style={{ color: data.budgetStatus.remaining < data.budgetStatus.budget * 0.1 ? '#D13438' : theme.palette.neutralPrimary }}>
            {formatCurrency(data.budgetStatus.remaining)}
          </span>
        </div>
      </div>

      <div className={classNames.divider} />

      {/* Earned Value */}
      <div>
        <Text style={{ fontSize: '12px', fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: '8px', display: 'block' }}>
          Earned Value Analysis
        </Text>
        <div className={classNames.evRow}>
          <div className={classNames.evItem}>
            <div className={classNames.evLabel}>SPI</div>
            <div className={classNames.evValue} style={{ color: data.earnedValue.spi >= 1 ? '#107C10' : '#D13438' }}>
              {data.earnedValue.spi.toFixed(2)}
            </div>
            <div style={{ fontSize: '9px', color: theme.palette.neutralTertiary }}>
              {data.earnedValue.spi >= 1 ? 'On Schedule' : 'Behind'}
            </div>
          </div>
          <div className={classNames.evItem}>
            <div className={classNames.evLabel}>CPI</div>
            <div className={classNames.evValue} style={{ color: data.earnedValue.cpi >= 1 ? '#107C10' : '#D13438' }}>
              {data.earnedValue.cpi.toFixed(2)}
            </div>
            <div style={{ fontSize: '9px', color: theme.palette.neutralTertiary }}>
              {data.earnedValue.cpi >= 1 ? 'On Budget' : 'Over Budget'}
            </div>
          </div>
          <div className={classNames.evItem}>
            <div className={classNames.evLabel}>Earned</div>
            <div className={classNames.evValue} style={{ fontSize: '13px', color: theme.palette.neutralPrimary }}>
              {formatCurrency(data.earnedValue.earnedValue)}
            </div>
          </div>
          <div className={classNames.evItem}>
            <div className={classNames.evLabel}>Actual</div>
            <div className={classNames.evValue} style={{ fontSize: '13px', color: data.earnedValue.actualCost > data.earnedValue.earnedValue ? '#D13438' : '#107C10' }}>
              {formatCurrency(data.earnedValue.actualCost)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ProgressDashboardCard;
