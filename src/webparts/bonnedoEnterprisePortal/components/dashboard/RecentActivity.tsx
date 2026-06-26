import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IRecentActivityProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IActivityItem {
    id: string;
    type: 'MR' | 'PR' | 'PO' | 'Payment';
    typeIcon: string;
    recordNumber: string;
    project: string;
    status: string;
    statusColor: string;
    createdDate: string;
}

const RecentActivity: React.FC<IRecentActivityProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [activities, setActivities] = React.useState<IActivityItem[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [isMobileView, setIsMobileView] = React.useState(false);

    React.useEffect(() => {
        const checkMobile = (): void => {
            setIsMobileView(window.innerWidth < 768);
        };
        checkMobile();
        window.addEventListener('resize', checkMobile);
        return () => window.removeEventListener('resize', checkMobile);
    }, []);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const [materialRequests, purchaseOrders, paymentRequests] = await Promise.all([
                    fetchListData(webUrl, 'PRC_Material_Request_Register'),
                    fetchListData(webUrl, 'PRC_Purchase_Order_Register'),
                    fetchListData(webUrl, 'FIN_Payment_Request_Register'),
                ]);

                const allActivities: IActivityItem[] = [];

                materialRequests.slice(0, 4).forEach((item: any) => {
                    allActivities.push({
                        id: `MR-${item.ID}`,
                        type: 'MR',
                        typeIcon: '📝',
                        recordNumber: item.Title || item.MR_Number || `MR-${item.ID}`,
                        project: item.Project || item.Project_Name || '-',
                        status: item.Approval_Status || item.Status || 'Pending',
                        statusColor: (item.Approval_Status === 'Approved' || item.Status === 'Approved')
                            ? theme.palette.green : theme.palette.orange,
                        createdDate: item.Created || item.Submitted_Date || new Date().toISOString(),
                    });
                });

                purchaseOrders.slice(0, 4).forEach((item: any) => {
                    allActivities.push({
                        id: `PO-${item.ID}`,
                        type: 'PO',
                        typeIcon: '🛒',
                        recordNumber: item.Title || item.PO_Number || `PO-${item.ID}`,
                        project: item.Project || item.Project_Name || '-',
                        status: item.Approval_Status || item.Status || 'Pending',
                        statusColor: (item.Approval_Status === 'Approved' || item.Status === 'Approved')
                            ? theme.palette.green : theme.palette.orange,
                        createdDate: item.Created || item.Submitted_Date || new Date().toISOString(),
                    });
                });

                paymentRequests.slice(0, 4).forEach((item: any) => {
                    allActivities.push({
                        id: `PAY-${item.ID}`,
                        type: 'Payment',
                        typeIcon: '💰',
                        recordNumber: item.Title || item.Payment_Number || `PAY-${item.ID}`,
                        project: item.Project || item.Project_Name || '-',
                        status: item.Approval_Status || 'Pending',
                        statusColor: item.Approval_Status === 'Approved' ? theme.palette.green : theme.palette.orange,
                        createdDate: item.Created || item.Submitted_Date || new Date().toISOString(),
                    });
                });

                allActivities.sort((a, b) =>
                    new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime()
                );
                setActivities(allActivities.slice(0, 10));
            } catch (error) {
                console.error('Error fetching recent activity:', error);
                setActivities([]);
            }
            setIsLoading(false);
        };
        fetchData();
    }, [spHttpClient, pageContext, refreshKey, theme]);

    const fetchListData = async (webUrl: string, listName: string): Promise<any[]> => {
        try {
            const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=10&$orderby=Created desc`;
            const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!response.ok) return [];
            const data = await response.json();
            return data.value || [];
        } catch { return []; }
    };

    const formatDate = (dateString: string): string => {
        try {
            const date = new Date(dateString);
            const now = new Date();
            const diffMs = now.getTime() - date.getTime();
            const diffMins = Math.floor(diffMs / 60000);
            const diffHours = Math.floor(diffMins / 60);
            const diffDays = Math.floor(diffHours / 24);
            if (diffMins < 1) return 'Just now';
            if (diffMins < 60) return `${diffMins}m ago`;
            if (diffHours < 24) return `${diffHours}h ago`;
            if (diffDays < 7) return `${diffDays}d ago`;
            return date.toLocaleDateString();
        } catch { return '-'; }
    };

    if (isLoading) {
        return React.createElement('div', {
            style: { height: '250px', display: 'flex', alignItems: 'center', justifyContent: 'center' }
        }, React.createElement('span', { style: { color: theme.palette.neutralSecondary } }, 'Loading...'));
    }

    if (activities.length === 0) {
        return React.createElement('div', {
            style: { height: '250px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }
        },
            React.createElement('span', { style: { fontSize: '32px', marginBottom: '8px' } }, '📋'),
            React.createElement('span', { style: { color: theme.palette.neutralSecondary } }, 'No recent activity')
        );
    }

    if (isMobileView) {
        return React.createElement('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '300px', overflowY: 'auto' } },
            activities.map((activity) =>
                React.createElement('div', {
                    key: activity.id,
                    style: {
                        padding: '10px 12px',
                        backgroundColor: theme.palette.neutralLighterAlt,
                        borderRadius: '8px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px',
                        borderLeft: `3px solid ${activity.statusColor}`,
                    }
                },
                    React.createElement('div', {
                        style: {
                            fontSize: '18px',
                            width: '32px',
                            height: '32px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            backgroundColor: theme.palette.white,
                            borderRadius: '6px',
                            flexShrink: 0,
                        }
                    }, activity.typeIcon),
                    React.createElement('div', { style: { flex: 1, minWidth: 0 } },
                        React.createElement('div', {
                            style: { display: 'flex', justifyContent: 'space-between', marginBottom: '2px' }
                        },
                            React.createElement('span', {
                                style: { fontSize: '12px', fontWeight: 600, color: theme.palette.neutralPrimary, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }
                            }, activity.recordNumber),
                            React.createElement('span', {
                                style: {
                                    fontSize: '10px', fontWeight: 500, color: activity.statusColor,
                                    backgroundColor: activity.statusColor + '20',
                                    padding: '2px 6px', borderRadius: '10px', flexShrink: 0, marginLeft: 6,
                                }
                            }, activity.status)
                        ),
                        React.createElement('div', {
                            style: { display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: theme.palette.neutralSecondary }
                        },
                            React.createElement('span', { style: { whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', flex: 1 } }, activity.project),
                            React.createElement('span', { style: { flexShrink: 0, marginLeft: 8 } }, formatDate(activity.createdDate))
                        )
                    )
                )
            )
        );
    }

    // Desktop: scrollable card layout (consistent, no table overflow)
    return React.createElement('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '320px', overflowY: 'auto' } },
        activities.map((activity) =>
            React.createElement('div', {
                key: activity.id,
                style: {
                    display: 'flex',
                    alignItems: 'center',
                    gap: '12px',
                    padding: '10px 12px',
                    borderRadius: '8px',
                    backgroundColor: theme.palette.neutralLighterAlt,
                    borderLeft: `3px solid ${activity.statusColor}`,
                    transition: 'background 0.15s ease',
                    cursor: 'default',
                }
            },
                React.createElement('div', {
                    style: {
                        fontSize: '18px',
                        width: '36px', height: '36px',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        backgroundColor: theme.palette.white,
                        borderRadius: '8px',
                        flexShrink: 0,
                    }
                }, activity.typeIcon),
                React.createElement('div', { style: { flex: 1, minWidth: 0 } },
                    React.createElement('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' } },
                        React.createElement('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: 0 } },
                            React.createElement('span', { style: { fontSize: '12px', fontWeight: 600, color: theme.palette.neutralSecondary, flexShrink: 0 } }, activity.type),
                            React.createElement('span', { style: { fontSize: '12px', fontWeight: 600, color: theme.palette.neutralPrimary, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, activity.recordNumber),
                        ),
                        React.createElement('div', { style: { display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 } },
                            React.createElement('span', {
                                style: {
                                    fontSize: '10px', fontWeight: 500, color: activity.statusColor,
                                    backgroundColor: activity.statusColor + '20',
                                    padding: '2px 8px', borderRadius: '10px',
                                }
                            }, activity.status),
                            React.createElement('span', { style: { fontSize: '11px', color: theme.palette.neutralTertiary } }, formatDate(activity.createdDate)),
                        )
                    ),
                    React.createElement('div', {
                        style: { fontSize: '12px', color: theme.palette.neutralSecondary, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }
                    }, activity.project !== '-' ? activity.project : 'No project')
                )
            )
        )
    );
};

export default RecentActivity;
