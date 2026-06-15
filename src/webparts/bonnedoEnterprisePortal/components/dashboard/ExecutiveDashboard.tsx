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

// KPI Metrics interface
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

    // Fetch all data
    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setMetrics(prev => ({ ...prev, isLoading: true }));
            try {
                const webUrl = pageContext.web.absoluteUrl;

                // Fetch data from all lists in parallel
                const [
                    purchaseOrders,
                    projects,
                    paymentRequests,
                    materialRequests,
                    purchaseRequisitions,
                ] = await Promise.all([
                    fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
                    fetchListData(webUrl, 'ENT_Project_Master'),
                    fetchListData(webUrl, 'FIN_Payment_Request_Register'),
                    fetchListData(webUrl, 'PRC_Material_Request_Register'),
                    fetchListData(webUrl, 'PRC_Purchase_Requisition_Register'),
                ]);

                // Calculate Total Procurement Value (sum of all PO amounts)
                const procurementValue = purchaseOrders.reduce((sum: number, item: any) =>
                    sum + (parseFloat(item.Amount || item.Total_Amount || item.PO_Amount || '0') || 0), 0
                );

                // Calculate Total Project Budget (sum of all project contract values)
                const projectBudget = projects.reduce((sum: number, item: any) =>
                    sum + (parseFloat(item.Contract_Value || item.Budget || item.Project_Budget || '0') || 0), 0
                );

                // Calculate Pending Approvals (from all request types)
                const pendingMR = materialRequests.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending'
                ).length;
                const pendingPR = purchaseRequisitions.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending'
                ).length;
                const pendingPO = purchaseOrders.filter((item: any) =>
                    item.Approval_Status === 'Pending' || item.Status === 'Pending'
                ).length;
                const pendingPayments = paymentRequests.filter((item: any) =>
                    item.Approval_Status === 'Pending'
                ).length;

                const totalPending = pendingMR + pendingPR + pendingPO + pendingPayments;

                // Calculate Payments This Month
                const now = new Date();
                const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();

                const paymentsThisMonth = paymentRequests
                    .filter((item: any) => {
                        const created = item.Created || item.Submitted_Date;
                        return created && created >= firstDayOfMonth && item.Approval_Status === 'Approved';
                    })
                    .reduce((sum: number, item: any) =>
                        sum + (parseFloat(item.Amount || '0') || 0), 0
                    );

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
        } catch {
            console.warn(`Failed to fetch ${listName}`);
            return [];
        }
    };

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
    };

    const formatCurrency = (amount: number): string => {
        if (amount >= 1000000000) return `N${(amount / 1000000000).toFixed(1)}B`;
        if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
        if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
        return `N${amount.toFixed(0)}`;
    };

    const classNames = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            backgroundColor: theme.palette.neutralLighterAlt,
            padding: '24px',
            gap: '24px',
        },
        headerSection: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
        },
        headerTitle: {
            display: 'flex',
            flexDirection: 'column',
        },
        welcomeText: {
            fontSize: '24px',
            fontWeight: 700,
            color: theme.palette.neutralPrimary,
        },
        subWelcomeText: {
            fontSize: '14px',
            color: theme.palette.neutralSecondary,
        },
        sectionTitle: {
            fontSize: '18px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
            marginBottom: '16px',
        },
        kpiGrid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
            gap: '16px',
        },
        header: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '16px',
        },
        title: {
            fontSize: '24px',
            fontWeight: 700,
            color: theme.palette.neutralPrimary,
        },
        subtitle: {
            fontSize: '14px',
            color: theme.palette.neutralSecondary,
        },
        kpiCard: {
            padding: '20px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            boxShadow: theme.effects.elevation4,
            display: 'flex',
            flexDirection: 'column',
            gap: '8px',
            transition: 'transform 0.2s ease, box-shadow 0.2s ease',
        },
        kpiIcon: {
            width: '48px',
            height: '48px',
            borderRadius: '12px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '24px',
        },
        kpiValue: {
            fontSize: '28px',
            fontWeight: 700,
            color: theme.palette.neutralPrimary,
        },
        kpiLabel: {
            fontSize: '14px',
            fontWeight: 500,
            color: theme.palette.neutralSecondary,
        },
        kpiDescription: {
            fontSize: '12px',
            color: theme.palette.neutralTertiary,
        },
        chartsGrid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))',
            gap: '20px',
        },
        chartCard: {
            padding: '20px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            boxShadow: theme.effects.elevation4,
            overflow: 'hidden',
        },
        chartTitle: {
            fontSize: '16px',
            fontWeight: 600,
            marginBottom: '16px',
            color: theme.palette.neutralPrimary,
        },
        fullWidthCard: {
            gridColumn: '1 / -1',
        },
    });

    // KPI Card component
    const KpiCard: React.FC<{
        icon: string;
        iconColor: string;
        iconBg: string;
        value: string;
        label: string;
        description: string;
        isLoading?: boolean;
    }> = ({ icon, iconColor, iconBg, value, label, description, isLoading }) => (
        <div className={classNames.kpiCard}>
            <div className={classNames.kpiIcon} style={{ backgroundColor: iconBg, color: iconColor }}>
                {icon}
            </div>
            {isLoading ? (
                <Shimmer width="60%" height="32px" />
            ) : (
                <div className={classNames.kpiValue}>{value}</div>
            )}
            <div className={classNames.kpiLabel}>{label}</div>
            <div className={classNames.kpiDescription}>{description}</div>
        </div>
    );

    return (
        <div className={classNames.root}>
            {/* Header */}
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" className={classNames.title}>
                        Executive Command Center
                    </Text>
                    <Text variant="medium" className={classNames.subtitle}>
                        Welcome back, {userDisplayName} - Here's your organizational overview
                    </Text>
                </div>
                <IconButton
                    iconProps={{ iconName: 'Refresh' }}
                    onClick={handleRefresh}
                    title="Refresh data"
                    ariaLabel="Refresh data"
                    style={{
                        backgroundColor: theme.palette.white,
                        borderRadius: '8px',
                        boxShadow: theme.effects.elevation4
                    }}
                />
            </div>

            {/* KPI Cards */}
            <div className={classNames.kpiGrid}>
                <KpiCard
                    icon="🛒"
                    iconColor={theme.palette.purple}
                    iconBg={theme.palette.purpleLight}
                    value={formatCurrency(metrics.totalProcurementValue)}
                    label="Total Procurement Value"
                    description="Sum of all purchase orders"
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="📊"
                    iconColor={theme.palette.blue}
                    iconBg={theme.palette.blueLight}
                    value={formatCurrency(metrics.totalProjectBudget)}
                    label="Total Project Budget"
                    description="Combined contract values"
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="⏳"
                    iconColor={theme.palette.orange}
                    iconBg={theme.palette.orangeLight}
                    value={metrics.pendingApprovals.toString()}
                    label="Pending Approvals"
                    description="MRs, PRs, POs & Payments"
                    isLoading={metrics.isLoading}
                />
                <KpiCard
                    icon="💰"
                    iconColor={theme.palette.green}
                    iconBg={theme.palette.greenLight}
                    value={formatCurrency(metrics.paymentsThisMonth)}
                    label="Payments This Month"
                    description="Approved payments in current month"
                    isLoading={metrics.isLoading}
                />
            </div>

            {/* Charts Row 1 */}
            <div className={classNames.chartsGrid}>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Top Vendor Spend
                    </Text>
                    <VendorSpendChart
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Procurement Pipeline
                    </Text>
                    <ProcurementTimeline
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
            </div>

            {/* Charts Row 2 */}
            <div className={classNames.chartsGrid}>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Budget Utilization by Project
                    </Text>
                    <BudgetUtilizationChart
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Recent Activity
                    </Text>
                    <RecentActivity
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
            </div>

            {/* Charts Row 3 - Additional Executive Insights */}
            <div className={classNames.chartsGrid}>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Payment Status Distribution
                    </Text>
                    <PaymentStatusChart
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Project Status Overview
                    </Text>
                    <ProjectStatusOverview
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
            </div>

            {/* Charts Row 4 - Operations & Master Data */}
            <div className={classNames.chartsGrid}>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Approval Metrics
                    </Text>
                    <ApprovalMetricsChart
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
                <div className={classNames.chartCard}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Material Categories
                    </Text>
                    <MaterialCategoryBreakdown
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
            </div>

            {/* Charts Row 5 - Procurement Performance */}
            <div className={classNames.chartsGrid}>
                <div className={`${classNames.chartCard} ${classNames.fullWidthCard}`}>
                    <Text variant="large" className={classNames.chartTitle}>
                        Procurement Performance by Stage
                    </Text>
                    <ProcurementPerformanceChart
                        spHttpClient={spHttpClient}
                        pageContext={pageContext}
                        refreshKey={refreshKey}
                    />
                </div>
            </div>
        </div>
    );
};

export default ExecutiveDashboard;
