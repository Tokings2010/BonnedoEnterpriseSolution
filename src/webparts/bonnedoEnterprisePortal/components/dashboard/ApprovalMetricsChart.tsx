import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IApprovalMetricsChartProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IApprovalData {
    type: string;
    pending: number;
    approved: number;
    rejected: number;
    avgDays: number;
}

const ApprovalMetricsChart: React.FC<IApprovalMetricsChartProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [approvalData, setApprovalData] = React.useState<IApprovalData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [totalPending, setTotalPending] = React.useState(0);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;

                // Fetch from all approval-related lists
                const [
                    materialRequests,
                    purchaseRequisitions,
                    purchaseOrders,
                    paymentRequests,
                ] = await Promise.all([
                    fetchListData(spHttpClient, webUrl, 'PRC_Material_Request_Register'),
                    fetchListData(spHttpClient, webUrl, 'PRC_Purchase_Requisition_Register'),
                    fetchListData(spHttpClient, webUrl, 'PRC_Purchase_Order_Register'),
                    fetchListData(spHttpClient, webUrl, 'FIN_Payment_Request_Register'),
                ]);

                const processListData = (items: any[], type: string): IApprovalData => {
                    const pending = items.filter((item: any) =>
                        (item.Status || item.Approval_Status || '').toLowerCase() === 'pending'
                    ).length;
                    const approved = items.filter((item: any) =>
                        (item.Status || item.Approval_Status || '').toLowerCase() === 'approved'
                    ).length;
                    const rejected = items.filter((item: any) =>
                        (item.Status || item.Approval_Status || '').toLowerCase() === 'rejected'
                    ).length;

                    return {
                        type,
                        pending,
                        approved,
                        rejected,
                        avgDays: Math.floor(Math.random() * 5) + 1, // Simulated avg days
                    };
                };

                const data: IApprovalData[] = [
                    processListData(materialRequests, 'Material Request'),
                    processListData(purchaseRequisitions, 'Purchase Requisition'),
                    processListData(purchaseOrders, 'Purchase Order'),
                    processListData(paymentRequests, 'Payment Request'),
                ];

                setApprovalData(data);
                setTotalPending(data.reduce((sum, d) => sum + d.pending, 0));
            } catch (error) {
                console.error('Error fetching approval metrics:', error);
                setApprovalData([]);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [refreshKey, pageContext]);

    const maxTotal = Math.max(...approvalData.map(d => d.pending + d.approved + d.rejected), 1);

    const colors = {
        pending: theme.palette.orange,
        approved: theme.palette.green,
        rejected: theme.palette.red,
    };

    if (isLoading) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                Loading approval metrics...
            </div>
        );
    }

    if (approvalData.length === 0) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                No approval data available
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {/* Summary Header */}
            <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                padding: '12px 16px',
                backgroundColor: totalPending > 0 ? `${theme.palette.orangeLight}` : theme.palette.neutralLighterAlt,
                borderRadius: '8px',
                borderLeft: `4px solid ${totalPending > 0 ? theme.palette.orange : theme.palette.green}`,
            }}>
                <div>
                    <div style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>
                        Pending Approvals
                    </div>
                    <div style={{ fontSize: '28px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                        {totalPending}
                    </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                    <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                        Needs Attention
                    </div>
                    <div style={{ fontSize: '14px', fontWeight: 600, color: theme.palette.orange }}>
                        {totalPending > 0 ? 'Action Required' : 'All Clear'}
                    </div>
                </div>
            </div>

            {/* Stacked Bar Chart */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {approvalData.map((item) => {
                    const total = item.pending + item.approved + item.rejected;
                    const pendingPercent = total > 0 ? (item.pending / maxTotal) * 100 : 0;
                    const approvedPercent = total > 0 ? (item.approved / maxTotal) * 100 : 0;
                    const rejectedPercent = total > 0 ? (item.rejected / maxTotal) * 100 : 0;

                    return (
                        <div key={item.type}>
                            <div style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                marginBottom: '4px',
                                fontSize: '13px',
                            }}>
                                <span style={{ fontWeight: 500, color: theme.palette.neutralPrimary }}>
                                    {item.type}
                                </span>
                                <span style={{ color: theme.palette.neutralSecondary }}>
                                    {item.pending > 0 && <span style={{ color: colors.pending }}>{item.pending} pending</span>}
                                    {item.pending > 0 && item.approved > 0 && ' • '}
                                    {item.approved > 0 && <span style={{ color: colors.approved }}>{item.approved} approved</span>}
                                    {(item.pending > 0 || item.approved > 0) && item.rejected > 0 && ' • '}
                                    {item.rejected > 0 && <span style={{ color: colors.rejected }}>{item.rejected} rejected</span>}
                                </span>
                            </div>
                            <div style={{
                                height: '24px',
                                display: 'flex',
                                borderRadius: '4px',
                                overflow: 'hidden',
                                backgroundColor: theme.palette.neutralLighter,
                            }}>
                                {item.pending > 0 && (
                                    <div style={{
                                        width: `${pendingPercent}%`,
                                        backgroundColor: colors.pending,
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                    }}>
                                        {item.pending > 0 && item.pending >= Math.floor(maxTotal * 0.1) && (
                                            <span style={{ color: 'white', fontSize: '11px', fontWeight: 600 }}>
                                                {item.pending}
                                            </span>
                                        )}
                                    </div>
                                )}
                                {item.approved > 0 && (
                                    <div style={{
                                        width: `${approvedPercent}%`,
                                        backgroundColor: colors.approved,
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                    }}>
                                        {item.approved > 0 && item.approved >= Math.floor(maxTotal * 0.1) && (
                                            <span style={{ color: 'white', fontSize: '11px', fontWeight: 600 }}>
                                                {item.approved}
                                            </span>
                                        )}
                                    </div>
                                )}
                                {item.rejected > 0 && (
                                    <div style={{
                                        width: `${rejectedPercent}%`,
                                        backgroundColor: colors.rejected,
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                    }}>
                                        {item.rejected > 0 && item.rejected >= Math.floor(maxTotal * 0.1) && (
                                            <span style={{ color: 'white', fontSize: '11px', fontWeight: 600 }}>
                                                {item.rejected}
                                            </span>
                                        )}
                                    </div>
                                )}
                            </div>
                        </div>
                    );
                })}
            </div>

            {/* Legend */}
            <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', paddingTop: '8px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <div style={{ width: '12px', height: '12px', borderRadius: '2px', backgroundColor: colors.pending }} />
                    <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>Pending</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <div style={{ width: '12px', height: '12px', borderRadius: '2px', backgroundColor: colors.approved }} />
                    <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>Approved</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <div style={{ width: '12px', height: '12px', borderRadius: '2px', backgroundColor: colors.rejected }} />
                    <span style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>Rejected</span>
                </div>
            </div>
        </div>
    );
};

// Helper function
async function fetchListData(spHttpClient: SPHttpClient, webUrl: string, listTitle: string): Promise<any[]> {
    try {
        const response = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$top=500`, SPHttpClient.configurations.v1);
        if (!response.ok) return [];
        const data = await response.json();
        return data.value || [];
    } catch {
        return [];
    }
}

export default ApprovalMetricsChart;
