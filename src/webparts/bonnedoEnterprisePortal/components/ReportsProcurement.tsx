import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    PrimaryButton,
    IconButton,
    Stack,
    Icon,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService, IListItem } from '../services/SharePointService';

// Simple bar chart component using HTML/CSS
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
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
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
                            title={`${item.label}: ${item.value}`}
                        />
                        <Text variant="small" style={{ marginTop: '8px', textAlign: 'center' }}>
                            {item.label.length > 10 ? item.label.substring(0, 10) + '...' : item.label}
                        </Text>
                        <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                            {item.value}
                        </Text>
                    </div>
                ))}
            </div>
        </div>
    );
};

// Simple pie chart component
const SimplePieChart: React.FC<{
    data: { label: string; value: number; color?: string }[];
    title: string;
    height?: number;
}> = ({ data, title, height = 250 }) => {
    const theme = getTheme();
    const total = data.reduce((sum, item) => sum + item.value, 0);
    const colors = [
        theme.palette.themePrimary,
        theme.palette.themeSecondary,
        '#107c10',
        '#ca5010',
        '#5c2d91',
        '#038387',
    ];

    let cumulativePercentage = 0;

    return (
        <div style={{
            padding: '16px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            border: `1px solid ${theme.palette.neutralLight}`,
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
        }}>
            <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                {title}
            </Text>
            <div style={{ display: 'flex', alignItems: 'center', gap: '24px' }}>
                <div style={{
                    width: '150px',
                    height: '150px',
                    borderRadius: '50%',
                    background: `conic-gradient(${data.map((item, index) => {
                        const percentage = (item.value / total) * 100;
                        const color = colors[index % colors.length];
                        const result = `${color} ${cumulativePercentage}% ${cumulativePercentage + percentage}%`;
                        cumulativePercentage += percentage;
                        return result;
                    }).join(', ')})`,
                }} />
                <div style={{ flex: 1 }}>
                    {data.map((item, index) => (
                        <div key={index} style={{ display: 'flex', alignItems: 'center', marginBottom: '8px' }}>
                            <div style={{
                                width: '12px',
                                height: '12px',
                                backgroundColor: colors[index % colors.length],
                                borderRadius: '2px',
                                marginRight: '8px',
                            }} />
                            <Text variant="small" style={{ flex: 1 }}>{item.label}</Text>
                            <Text variant="small" style={{ fontWeight: 600 }}>
                                {item.value} ({((item.value / total) * 100).toFixed(1)}%)
                            </Text>
                        </div>
                    ))}
                </div>
            </div>
        </div>
    );
};

// Export to CSV function
const exportToCSV = (data: IListItem[], filename: string): void => {
    if (data.length === 0) return;

    const headers = Object.keys(data[0]).filter(key => key !== 'ID');
    const csvContent = [
        headers.join(','),
        ...data.map(row =>
            headers.map(header => {
                const value = row[header];
                // Handle objects and arrays
                if (typeof value === 'object') {
                    return `"${JSON.stringify(value).replace(/"/g, '""')}"`;
                }
                // Escape quotes and wrap in quotes
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

// Export to Excel function - placeholder for future use
// eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-empty-function
const exportToExcel = (_data: IListItem[], _filename: string): void => {
    // Reserved for xlsx library integration
};

// Procurement Reports Component
interface IReportsProcurementProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const ReportsProcurement: React.FC<IReportsProcurementProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [loading, setLoading] = React.useState(true);
    const [materialRequests, setMaterialRequests] = React.useState<IListItem[]>([]);
    const [purchaseRequisitions, setPurchaseRequisitions] = React.useState<IListItem[]>([]);
    const [purchaseOrders, setPurchaseOrders] = React.useState<IListItem[]>([]);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    React.useEffect(() => {
        const fetchData = async () => {
            setLoading(true);
            try {
                const [mr, pr, po] = await Promise.all([
                    sharePointService.getListData('PRC_Material_Request_Register', undefined, 500),
                    sharePointService.getListData('PRC_Purchase_Requisition_Register', undefined, 500),
                    sharePointService.getListData('PRC_Purchase_Order_Register', undefined, 500),
                ]);
                setMaterialRequests(mr);
                setPurchaseRequisitions(pr);
                setPurchaseOrders(po);
            } catch (error) {
                console.error('Error fetching report data:', error);
            }
            setLoading(false);
        };
        fetchData();
    }, [sharePointService]);

    // Process data for charts
    const mrByProject = React.useMemo(() => {
        const counts: Record<string, number> = {};
        materialRequests.forEach(mr => {
            const projectCode = mr.Project_Code || 'Unknown';
            counts[projectCode] = (counts[projectCode] || 0) + 1;
        });
        return Object.entries(counts)
            .map(([label, value]) => ({ label, value }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 8);
    }, [materialRequests]);

    const prStatusData = React.useMemo(() => {
        const approved = purchaseRequisitions.filter(pr => pr.Approval_Status === 'Approved').length;
        const pending = purchaseRequisitions.filter(pr => pr.Approval_Status === 'Pending').length;
        const rejected = purchaseRequisitions.filter(pr => pr.Approval_Status === 'Rejected').length;
        return [
            { label: 'Approved', value: approved, color: '#107c10' },
            { label: 'Pending', value: pending, color: '#ca5010' },
            { label: 'Rejected', value: rejected, color: '#a80000' },
        ];
    }, [purchaseRequisitions]);

    const poByVendor = React.useMemo(() => {
        const totals: Record<string, number> = {};
        purchaseOrders.forEach(po => {
            const vendor = po.Vendor || 'Unknown';
            totals[vendor] = (totals[vendor] || 0) + (po.Amount || po.TotalAmount || 0);
        });
        return Object.entries(totals)
            .map(([label, value]) => ({ label, value: Math.round(value) }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 6);
    }, [purchaseOrders]);

    const handleExportMR = (): void => exportToCSV(materialRequests, 'MaterialRequests');
    const handleExportPR = (): void => exportToCSV(purchaseRequisitions, 'PurchaseRequisitions');
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
                        Procurement Reports
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Analyze procurement data and trends
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
                    <span style={{ marginLeft: '8px' }}>Loading report data...</span>
                </div>
            ) : (
                <>
                    <div className={classNames.chartsGrid}>
                        <SimpleBarChart
                            title="Material Requests by Project"
                            data={mrByProject}
                        />
                        <SimplePieChart
                            title="PR Approved vs Pending"
                            data={prStatusData}
                        />
                        <SimpleBarChart
                            title="PO Value by Vendor (₦)"
                            data={poByVendor}
                            height={200}
                        />
                    </div>

                    {/* Export Section */}
                    <div className={classNames.exportSection}>
                        <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: '16px' }}>
                            Export Data
                        </Text>
                        <Stack horizontal tokens={{ childrenGap: '12px' }}>
                            <PrimaryButton
                                text="Export MR to Excel"
                                iconProps={{ iconName: 'ExcelLogo' }}
                                onClick={handleExportMR}
                            />
                            <PrimaryButton
                                text="Export PR to Excel"
                                iconProps={{ iconName: 'ExcelLogo' }}
                                onClick={handleExportPR}
                            />
                            <PrimaryButton
                                text="Export PO to Excel"
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

export default ReportsProcurement;
