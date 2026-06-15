import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IVendorSpendChartProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface IVendorData {
    vendorName: string;
    totalSpend: number;
    color: string;
}

const VendorSpendChart: React.FC<IVendorSpendChartProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [vendorData, setVendorData] = React.useState<IVendorData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    const colors = [
        theme.palette.blue,
        theme.palette.green,
        theme.palette.orange,
        theme.palette.purple,
        theme.palette.red,
    ];

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const url = `${webUrl}/_api/web/lists/getByTitle('PRC_Purchase_Order_Register')/items?$top=500&$select=*,Vendor/Title&$expand=Vendor`;

                const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
                if (!response.ok) {
                    setVendorData([]);
                    setIsLoading(false);
                    return;
                }

                const data = await response.json();
                const items = data.value || [];

                // Group by vendor and calculate total spend
                const vendorMap = new Map<string, number>();

                items.forEach((item: any) => {
                    const vendorName = item.Vendor?.Title || item.Vendor_Name || item.Vendor || 'Unknown';
                    const amount = parseFloat(item.Amount || item.Total_Amount || item.PO_Amount || '0') || 0;

                    if (vendorMap.has(vendorName)) {
                        vendorMap.set(vendorName, vendorMap.get(vendorName)! + amount);
                    } else {
                        vendorMap.set(vendorName, amount);
                    }
                });

                // Convert to array and sort by spend, take top 5
                const sortedVendors = Array.from(vendorMap.entries())
                    .map(([name, spend], index) => ({
                        vendorName: name,
                        totalSpend: spend,
                        color: colors[index % colors.length],
                    }))
                    .sort((a, b) => b.totalSpend - a.totalSpend)
                    .slice(0, 5);

                setVendorData(sortedVendors);
            } catch (error) {
                console.error('Error fetching vendor data:', error);
                setVendorData([]);
            }
            setIsLoading(false);
        };

        fetchData();
    }, [spHttpClient, pageContext, refreshKey]);

    const maxSpend = Math.max(...vendorData.map(v => v.totalSpend), 1);

    const formatCurrency = (amount: number): string => {
        if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
        if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
        return `N${amount.toFixed(0)}`;
    };

    if (isLoading) {
        return (
            <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <span style={{ color: theme.palette.neutralSecondary }}>Loading...</span>
            </div>
        );
    }

    if (vendorData.length === 0) {
        return (
            <div style={{
                height: '200px',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center'
            }}>
                <span style={{ fontSize: '32px', marginBottom: '8px' }}>📊</span>
                <span style={{ color: theme.palette.neutralSecondary }}>No vendor data available</span>
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
            {vendorData.map((vendor, index) => (
                <div key={vendor.vendorName} style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{ width: '80px', fontSize: '12px', fontWeight: 500, color: theme.palette.neutralPrimary, flexShrink: 0 }}>
                        {vendor.vendorName.length > 10 ? vendor.vendorName.substring(0, 10) + '...' : vendor.vendorName}
                    </div>
                    <div style={{ flex: 1, position: 'relative' }}>
                        <div
                            style={{
                                height: '24px',
                                width: `${(vendor.totalSpend / maxSpend) * 100}%`,
                                backgroundColor: vendor.color,
                                borderRadius: '4px',
                                minWidth: '20px',
                            }}
                        />
                    </div>
                    <div style={{ width: '70px', textAlign: 'right', fontSize: '12px', fontWeight: 600, color: theme.palette.neutralPrimary }}>
                        {formatCurrency(vendor.totalSpend)}
                    </div>
                </div>
            ))}
        </div>
    );
};

export default VendorSpendChart;
