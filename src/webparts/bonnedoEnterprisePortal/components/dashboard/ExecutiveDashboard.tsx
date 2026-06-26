import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    Shimmer,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import VendorSpendChart from './VendorSpendChart';
import ProcurementTimeline from './ProcurementTimeline';
import BudgetUtilizationChart from './BudgetUtilizationChart';
import RecentActivity from './RecentActivity';
import ProcurementPerformanceChart from './ProcurementPerformanceChart';
import PaymentStatusChart from './PaymentStatusChart';
import ProjectStatusOverview from './ProjectStatusOverview';
import ApprovalMetricsChart from './ApprovalMetricsChart';
import MaterialCategoryBreakdown from './MaterialCategoryBreakdown';

export interface IExecutiveDashboardProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    userDisplayName: string;
}

interface IKpiMetrics {
    totalProcurementValue: number;
    totalProjectBudget: number;
    pendingApprovals: number;
    paymentsThisMonth: number;
    isLoading: boolean;
}

const ExecutiveDashboard: React.FC<IExecutiveDashboardProps> = ({
    spHttpClient,
    pageContext,
    userDisplayName,
}) => {
    const theme = getTheme();
    const [refreshKey, setRefreshKey] = React.useState(0);
    const [metrics, setMetrics] = React.useState<IKpiMetrics>({
        totalProcurementValue: 0,
        totalProjectBudget: 0,
        pendingApprovals: 0,
        paymentsThisMonth: 0,
        isLoading: true,
    });

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setMetrics(prev => ({ ...prev, isLoading: true }));
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const [purchaseOrders, projects, paymentRequests, materialRequests, purchaseRequisitions] =
                    await Promise.all([
                        fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
                        fetchListData(webUrl, 'ENT_Project_Master'),
                        fetchListData(webUrl, 'FIN_Payment_Request_Register'),
                        fetchListData(webUrl, 'PRC_Material_Request_Register'),
                        fetchListData(webUrl, 'PRC_Purchase_Requisition_Register'),
                    ]);

                const procurementValue = purchaseOrders.reduce((sum: number, item: any) =>
                    sum + (parseFloat(item.Amount || item.Total_Amount || item.PO_Amount || '0') || 0), 0);
                const projectBudget = projects.reduce((sum: number, item: any) =>
                    sum + (parseFloat(item.Contract_Value || item.Budget || item.Project_Budget || '0') || 0), 0);
                const pendingMR = materialRequests.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending').length;
                const pendingPR = purchaseRequisitions.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending').length;
                const pendingPO = purchaseOrders.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending').length;
                const pendingPayments = paymentRequests.filter((item: any) =>
                    item.Approval_Status === 'Pending').length;
                const totalPending = pendingMR + pendingPR + pendingPO + pendingPayments;

                const now = new Date();
                const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();
                const paymentsThisMonth = paymentRequests
                    .filter((item: any) => {
                        const created = item.Created || item.Submitted_Date;
                        return created && created >= firstDayOfMonth && item.Approval_Status === 'Approved';
                    })
                    .reduce((sum: number, item: any) => sum + (parseFloat(item.Amount || '0') || 0), 0);

                setMetrics({
                    totalProcurementValue: procurementValue,
                    totalProjectBudget: projectBudget,
                    pendingApprovals: totalPending,
                    paymentsThisMonth: paymentsThisMonth,
                    isLoading: false,
                });
            } catch (error) {
                console.error('Error fetching Executive Dashboard data:', error);
                setMetrics(prev => ({ ...prev, isLoading: false }));
            }
        };
        fetchData();
    }, [pageContext, refreshKey]);

    const fetchListData = async (webUrl: string, listName: string): Promise<any[]> => {
        try {
            const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=500&$select=*`;
            const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!response.ok) return [];
            const data = await response.json();
            return data.value || [];
        } catch { console.warn(`Failed to fetch ${listName}`); return []; }
    };

    const handleRefresh = (): void => setRefreshKey(prev => prev + 1);

    const formatCurrency = (amount: number): string => {
        if (amount >= 1000000000) return `N${(amount / 1000000000).toFixed(1)}B`;
        if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
        if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
        return `N${amount.toFixed(0)}`;
    };

    const c = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            gap: '24px',
            fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
        },
        headerBar: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'flex-start',
        },
        headerLeft: { display: 'flex', flexDirection: 'column', gap: '4px' },
        sectionTitle: {
            fontSize: '16px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
            marginBottom: '16px',
        },
        // KPI row – fixed 4 columns
        kpiGrid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(4, 1fr)',
            gap: '16px',
            '@media (max-width: 900px)': { gridTemplateColumns: 'repeat(2, 1fr)' },
            '@media (max-width: 500px)': { gridTemplateColumns: '1fr' },
        },
        kpiCard: {
            padding: '20px',
            backgroundColor: theme.palette.white,
            borderRadius: '12px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
            border: `1px solid ${theme.palette.neutralLighter}`,
            display: 'flex',
            flexDirection: 'column',
            gap: '8px',
            transition: 'transform 0.2s ease, box-shadow 0.2s ease',
            selectors: {
                ':hover': { transform: 'translateY(-2px)', boxShadow: '0 4px 12px rgba(0,0,0,0.08)' },
            },
        },
        kpiIconRow: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
        },
        kpiIconBox: {
            width: '44px',
            height: '44px',
            borderRadius: '12px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '22px',
        },
        kpiTrend: {
            fontSize: '12px',
            fontWeight: 600,
            padding: '2px 8px',
            borderRadius: '20px',
        },
        kpiValue: { fontSize: '28px', fontWeight: 700, color: theme.palette.neutralPrimary },
        kpiLabel: { fontSize: '13px', fontWeight: 500, color: theme.palette.neutralSecondary },
        kpiDescription: { fontSize: '11px', color: theme.palette.neutralTertiary },
        // Charts grid – exactly 2 columns
        chartsRow: {
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: '20px',
            '@media (max-width: 1000px)': { gridTemplateColumns: '1fr' },
        },
        chartCard: {
            padding: '20px',
            backgroundColor: theme.palette.white,
            borderRadius: '12px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
            border: `1px solid ${theme.palette.neutralLighter}`,
            overflow: 'hidden',
        },
        chartTitle: {
            fontSize: '15px',
            fontWeight: 600,
            marginBottom: '16px',
            color: theme.palette.neutralPrimary,
        },
        fullWidthCard: { gridColumn: '1 / -1' },
    });

    const KpiCard: React.FC<{
        icon: string; iconColor: string; iconBg: string;
        value: string; label: string; description: string;
        trend?: string; trendColor?: string;
        isLoading?: boolean;
    }> = ({ icon, iconColor, iconBg, value, label, description, trend, trendColor, isLoading }) => (
        <div className={c.kpiCard}>
            <div className={c.kpiIconRow}>
                <div className={c.kpiIconBox} style={{ backgroundColor: iconBg, color: iconColor }}>{icon}</div>
                {trend && <div className={c.kpiTrend} style={{ background: `${trendColor || theme.palette.green}15`, color: trendColor || theme.palette.green }}>{trend}</div>}
            </div>
            {isLoading ? <Shimmer width="60%" height="32px" /> : <div className={c.kpiValue}>{value}</div>}
            <div className={c.kpiLabel}>{label}</div>
            <div className={c.kpiDescription}>{description}</div>
        </div>
    );

    return (
        <div className={c.root}>
            {/* Header */}
            <div className={c.headerBar}>
                <div className={c.headerLeft}>
                    <Text variant="xxLarge" style={{ fontSize: '26px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                        Executive Command Center
                    </Text>
                    <Text variant="medium" style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>
                        Welcome back, {userDisplayName} — Here's your organizational overview
                    </Text>
                </div>
                <IconButton
                    iconProps={{ iconName: 'Refresh' }}
                    onClick={handleRefresh}
                    title="Refresh data"
                    ariaLabel="Refresh data"
                    styles={{ root: { background: theme.palette.white, borderRadius: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.06)', border: `1px solid ${theme.palette.neutralLighter}` } }}
                />
            </div>

            {/* KPI Cards */}
            <div className={c.kpiGrid}>
                <KpiCard
                    icon="🛒" iconColor={theme.palette.purple} iconBg={theme.palette.purpleLight}
                    value={formatCurrency(metrics.totalProcurementValue)}
                    label="Total Procurement Value" description="Sum of all purchase orders"
                    trend="+12%" trendColor={theme.palette.green}
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="📊" iconColor={theme.palette.blue} iconBg={theme.palette.blueLight}
                    value={formatCurrency(metrics.totalProjectBudget)}
                    label="Total Project Budget" description="Combined contract values"
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="⏳" iconColor={theme.palette.orange} iconBg={theme.palette.orangeLight}
                    value={metrics.pendingApprovals.toString()}
                    label="Pending Approvals" description="MRs, PRs, POs & Payments"
                    trend={metrics.pendingApprovals > 0 ? `${metrics.pendingApprovals} pending` : 'All clear'}
                    trendColor={metrics.pendingApprovals > 0 ? theme.palette.orange : theme.palette.green}
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="💰" iconColor={theme.palette.green} iconBg={theme.palette.greenLight}
                    value={formatCurrency(metrics.paymentsThisMonth)}
                    label="Payments This Month" description="Approved payments in current month"
                    isLoading={metrics.isLoading}
                />
            </div>

            {/* Charts Row 1 */}
            <div className={c.chartsRow}>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Top Vendor Spend</div>
                    <VendorSpendChart spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Procurement Pipeline</div>
                    <ProcurementTimeline spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
            </div>

            {/* Charts Row 2 */}
            <div className={c.chartsRow}>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Budget Utilization by Project</div>
                    <BudgetUtilizationChart spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Recent Activity</div>
                    <RecentActivity spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
            </div>

            {/* Charts Row 3 */}
            <div className={c.chartsRow}>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Payment Status Distribution</div>
                    <PaymentStatusChart spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Project Status Overview</div>
                    <ProjectStatusOverview spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
            </div>

            {/* Charts Row 4 */}
            <div className={c.chartsRow}>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Approval Metrics</div>
                    <ApprovalMetricsChart spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
                <div className={c.chartCard}>
                    <div className={c.chartTitle}>Material Categories</div>
                    <MaterialCategoryBreakdown spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
            </div>

            {/* Full width row */}
            <div className={c.chartsRow}>
                <div className={`${c.chartCard} ${c.fullWidthCard}`}>
                    <div className={c.chartTitle}>Procurement Performance by Stage</div>
                    <ProcurementPerformanceChart spHttpClient={spHttpClient} pageContext={pageContext} refreshKey={refreshKey} />
                </div>
            </div>
        </div>
    );
};

export default ExecutiveDashboard;
