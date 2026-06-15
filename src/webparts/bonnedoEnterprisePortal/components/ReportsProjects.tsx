import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    PrimaryButton,
    Stack,
    Icon,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

// Simple horizontal bar chart for budget
const BudgetBarChart: React.FC<{
    data: { label: string; used: number; total: number }[];
    title: string;
    height?: number;
}> = ({ data, title, height = 300 }) => {
    const theme = getTheme();

    return (
        <div style={{
            padding: '16px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            border: `1px solid ${theme.palette.neutralLight}`,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            boxShadow: '0 1px 3px rgba(0,0,0,0.12)' as any,
        }}>
            <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                {title}
            </Text>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {data.map((item, index) => {
                    const percentage = (item.used / item.total) * 100;
                    const getColor = () => {
                        if (percentage > 90) return '#a80000';
                        if (percentage > 70) return '#ca5010';
                        return '#107c10';
                    };
                    return (
                        <div key={index}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                                <Text variant="small" style={{ fontWeight: 500 }}>
                                    {item.label}
                                </Text>
                                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                                    ₦{item.used.toLocaleString()} / ₦{item.total.toLocaleString()} ({percentage.toFixed(1)}%)
                                </Text>
                            </div>
                            <div style={{
                                height: '20px',
                                backgroundColor: theme.palette.neutralLighter,
                                borderRadius: '10px',
                                overflow: 'hidden',
                            }}>
                                <div style={{
                                    width: `${Math.min(percentage, 100)}%`,
                                    height: '100%',
                                    backgroundColor: getColor(),
                                    borderRadius: '10px',
                                    transition: 'width 0.3s ease',
                                }} />
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
    );
};

// Count bar chart
const CountBarChart: React.FC<{
    data: { label: string; value: number }[];
    title: string;
    height?: number;
}> = ({ data, title, height = 250 }) => {
    const theme = getTheme();
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const maxValue = Math.max(...data.map(d => d.value), 1);

    return (
        <div style={{
            padding: '16px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            border: `1px solid ${theme.palette.neutralLight}`,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            boxShadow: '0 1px 3px rgba(0,0,0,0.12)' as any,
        }}>
            <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                {title}
            </Text>
            <div style={{ display: 'flex', alignItems: 'flex-end', gap: '8px', height: `${height}px` }}>
                {data.map((item, index) => (
                    <div key={index} style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div
                            style={{
                                width: '100%',
                                height: `${(item.value / maxValue) * (height - 40)}px`,
                                backgroundColor: theme.palette.themePrimary,
                                borderRadius: '4px 4px 0 0',
                                minHeight: '4px',
                            }}
                            title={`${item.label}: ${item.value}`}
                        />
                        <Text variant="small" style={{ marginTop: '8px', textAlign: 'center' }}>
                            {item.label.length > 10 ? item.label.substring(0, 10) + '...' : item.label}
                        </Text>
                        <Text variant="small" style={{ fontWeight: 600 }}>
                            {item.value}
                        </Text>
                    </div>
                ))}
            </div>
        </div>
    );
};

// Export to CSV
const exportToCSV = (data: IListItem[], filename: string): void => {
    if (data.length === 0) return;

    const headers = Object.keys(data[0]).filter(key => key !== 'ID');
    const csvContent = [
        headers.join(','),
        ...data.map(row =>
            headers.map(header => {
                const value = row[header];
                if (typeof value === 'object') {
                    return `"${JSON.stringify(value).replace(/"/g, '""')}"`;
                }
                return `"${String(value || '').replace(/"/g, '""')}"`;
            }).join(',')
        )
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
};

// Project Reports Component
interface IReportsProjectsProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const ReportsProjects: React.FC<IReportsProjectsProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [loading, setLoading] = React.useState(true);
    const [projects, setProjects] = React.useState<IListItem[]>([]);
    const [purchaseOrders, setPurchaseOrders] = React.useState<IListItem[]>([]);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    React.useEffect(() => {
        const fetchData = async () => {
            setLoading(true);
            try {
                const [proj, po] = await Promise.all([
                    sharePointService.getListData('ENT_Project_Master', undefined, 100),
                    sharePointService.getListData('PRC_Purchase_Order_Register', undefined, 500),
                ]);
                setProjects(proj);
                setPurchaseOrders(po);
            } catch (error) {
                console.error('Error fetching project data:', error);
            }
            setLoading(false);
        };
        fetchData();
    }, [sharePointService]);

    // Budget Utilization
    const budgetUtilization = React.useMemo(() => {
        return projects.map(project => {
            const projectCode = project.Project_Code || project.Title;
            const total = project.Contract_Value || project.Budget || 0;
            const used = purchaseOrders
                .filter(po => po.Project_Code === projectCode)
                .reduce((sum, po) => sum + (po.Amount || po.TotalAmount || 0), 0);
            return {
                label: projectCode,
                used: Math.round(used),
                total: Math.round(total),
            };
        }).filter(p => p.total > 0).slice(0, 8);
    }, [projects, purchaseOrders]);

    // Procurement Count by Project
    const procurementCount = React.useMemo(() => {
        const counts: Record<string, number> = {};
        purchaseOrders.forEach(po => {
            const projectCode = po.Project_Code || 'Unknown';
            counts[projectCode] = (counts[projectCode] || 0) + 1;
        });
        return Object.entries(counts)
            .map(([label, value]) => ({ label, value }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 8);
    }, [purchaseOrders]);

    const handleExportProjects = (): void => exportToCSV(projects, 'Projects');
    const handleExportPO = (): void => exportToCSV(purchaseOrders, 'PurchaseOrders');

    const classNames = mergeStyleSets({
        container: { padding: '20px' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' },
        chartsGrid: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))',
            gap: '24px',
            marginBottom: '24px',
        },
        exportSection: {
            marginTop: '24px',
            padding: '16px',
            backgroundColor: theme.palette.neutralLighterAlt,
            borderRadius: '8px',
        },
        loadingContainer: {
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            height: '200px',
            color: theme.palette.neutralSecondary,
        },
    });

    return (
        <div className={classNames.container}>
            {/* Tab Header */}
            <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Project Reports
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Analyze project progress and metrics
                    </Text>
                </div>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={onRefresh}
                        title="Refresh data"
                        ariaLabel="Refresh data"
                    />
                </div>
            </div>

            {loading ? (
                <div className={classNames.loadingContainer}>
                    <Icon iconName="Sync" style={{ fontSize: '24px', animation: 'spin 1s linear infinite' }} />
                    <span style={{ marginLeft: '8px' }}>Loading project data...</span>
                </div>
            ) : (
                <>
                    <div className={classNames.chartsGrid}>
                        <BudgetBarChart
                            title="Project Budget Utilization (₦)"
                            data={budgetUtilization}
                        />
                        <CountBarChart
                            title="Procurement Count by Project"
                            data={procurementCount}
                        />
                    </div>

                    <div className={classNames.exportSection}>
                        <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                            Export Data
                        </Text>
                        <Stack horizontal tokens={{ childrenGap: '12px' }}>
                            <PrimaryButton
                                text="Export Projects to Excel"
                                iconProps={{ iconName: 'ExcelLogo' }}
                                onClick={handleExportProjects}
                            />
                            <PrimaryButton
                                text="Export POs to Excel"
                                iconProps={{ iconName: 'ExcelLogo' }}
                                onClick={handleExportPO}
                            />
                        </Stack>
                    </div>
                </>
            )}
        </div>
    );
};

export default ReportsProjects;
