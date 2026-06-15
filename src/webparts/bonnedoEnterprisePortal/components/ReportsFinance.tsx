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

// Simple bar chart component
const SimpleBarChart: React.FC<{
    data: { label: string; value: number; color?: string }[];
    title: string;
    height?: number;
}> = ({ data, title, height = 250 }) => {
    const theme = getTheme();
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
                                backgroundColor: item.color || theme.palette.themePrimary,
                                borderRadius: '4px 4px 0 0',
                                minHeight: '4px',
                                transition: 'height 0.3s ease',
                            }}
                            title={`${item.label}: ₦${item.value.toLocaleString()}`}
                        />
                        <Text variant="small" style={{ marginTop: '8px', textAlign: 'center' }}>
                            {item.label.length > 10 ? item.label.substring(0, 10) + '...' : item.label}
                        </Text>
                        <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                            ₦{item.value.toLocaleString()}
                        </Text>
                    </div>
                ))}
            </div>
        </div>
    );
};

// Line chart component for trends
const SimpleLineChart: React.FC<{
    data: { label: string; value: number }[];
    title: string;
    height?: number;
}> = ({ data, title, height = 250 }) => {
    const theme = getTheme();
    const maxValue = Math.max(...data.map(d => d.value), 1);
    const minValue = Math.min(...data.map(d => d.value), 0);
    const range = maxValue - minValue || 1;

    const points = data.map((item, index) => {
        const x = (index / (data.length - 1)) * 100;
        const y = height - 30 - ((item.value - minValue) / range) * (height - 60);
        return { x, y, ...item };
    });

    const pathD = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');

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
            <svg width="100%" height={height} style={{ overflow: 'visible' }}>
                {/* Grid lines */}
                {[0, 0.25, 0.5, 0.75, 1].map((ratio, i) => (
                    <line
                        key={i}
                        x1="0"
                        y1={`${30 + (height - 60) * ratio}`}
                        x2="100%"
                        y2={`${30 + (height - 60) * ratio}`}
                        stroke={theme.palette.neutralLight}
                        strokeDasharray="4"
                    />
                ))}
                {/* Line path */}
                <path
                    d={pathD}
                    fill="none"
                    stroke={theme.palette.themePrimary}
                    strokeWidth="3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                />
                {/* Points */}
                {points.map((p, i) => (
                    <circle
                        key={i}
                        cx={`${p.x}%`}
                        cy={p.y}
                        r="5"
                        fill={theme.palette.themePrimary}
                        stroke={theme.palette.white}
                        strokeWidth="2"
                    />
                ))}
            </svg>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '8px' }}>
                {data.map((item, i) => (
                    <Text key={i} variant="small" style={{ color: theme.palette.neutralSecondary }}>
                        {item.label}
                    </Text>
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

// Finance Reports Component
interface IReportsFinanceProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const ReportsFinance: React.FC<IReportsFinanceProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [loading, setLoading] = React.useState(true);
    const [payments, setPayments] = React.useState<IListItem[]>([]);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    React.useEffect(() => {
        const fetchData = async () => {
            setLoading(true);
            try {
                const data = await sharePointService.getListData('FIN_Payment_Request_Register', undefined, 500);
                setPayments(data);
            } catch (error) {
                console.error('Error fetching finance data:', error);
            }
            setLoading(false);
        };
        fetchData();
    }, [sharePointService]);

    // Payments by Project
    const paymentsByProject = React.useMemo(() => {
        const totals: Record<string, number> = {};
        payments.forEach(p => {
            const projectCode = p.Project_Code || 'Unknown';
            if (p.Approval_Status === 'Approved' || p.Payment_Status === 'Paid') {
                totals[projectCode] = (totals[projectCode] || 0) + (p.Amount || 0);
            }
        });
        return Object.entries(totals)
            .map(([label, value]) => ({ label, value: Math.round(value) }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 6);
    }, [payments]);

    // Payments by Vendor
    const paymentsByVendor = React.useMemo(() => {
        const totals: Record<string, number> = {};
        payments.forEach(p => {
            const vendor = p.Vendor || 'Unknown';
            if (p.Approval_Status === 'Approved' || p.Payment_Status === 'Paid') {
                totals[vendor] = (totals[vendor] || 0) + (p.Amount || 0);
            }
        });
        return Object.entries(totals)
            .map(([label, value]) => ({ label, value: Math.round(value) }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 6);
    }, [payments]);

    // Monthly Payment Trend
    const monthlyTrend = React.useMemo(() => {
        const months: Record<string, number> = {};
        payments.forEach(p => {
            if (p.Approval_Status === 'Approved' || p.Payment_Status === 'Paid') {
                const date = p.Payment_Date || p.Modified || p.Created;
                if (date) {
                    const month = new Date(date).toLocaleString('en-US', { month: 'short', year: '2-digit' });
                    months[month] = (months[month] || 0) + (p.Amount || 0);
                }
            }
        });
        return Object.entries(months)
            .map(([label, value]) => ({ label, value: Math.round(value) }))
            .slice(-6);
    }, [payments]);

    const handleExport = (): void => exportToCSV(payments, 'FinancePayments');

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
                        Finance Reports
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Analyze financial performance and metrics
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
                    <span style={{ marginLeft: '8px' }}>Loading finance data...</span>
                </div>
            ) : (
                <>
                    <div className={classNames.chartsGrid}>
                        <SimpleBarChart
                            title="Payments by Project (₦)"
                            data={paymentsByProject}
                        />
                        <SimpleBarChart
                            title="Payments by Vendor (₦)"
                            data={paymentsByVendor}
                        />
                        <SimpleLineChart
                            title="Monthly Payment Trend (₦)"
                            data={monthlyTrend}
                        />
                    </div>

                    <div className={classNames.exportSection}>
                        <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                            Export Data
                        </Text>
                        <Stack horizontal tokens={{ childrenGap: '12px' }}>
                            <PrimaryButton
                                text="Export to Excel"
                                iconProps={{ iconName: 'ExcelLogo' }}
                                onClick={handleExport}
                            />
                        </Stack>
                    </div>
                </>
            )}
        </div>
    );
};

export default ReportsFinance;
