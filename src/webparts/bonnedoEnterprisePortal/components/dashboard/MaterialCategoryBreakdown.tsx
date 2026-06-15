import * as React from 'react';
import { getTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

interface IMaterialCategoryBreakdownProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    refreshKey: number;
}

interface ICategoryData {
    category: string;
    count: number;
    totalValue: number;
}

const MaterialCategoryBreakdown: React.FC<IMaterialCategoryBreakdownProps> = ({
    spHttpClient,
    pageContext,
    refreshKey,
}) => {
    const theme = getTheme();
    const [categoryData, setCategoryData] = React.useState<ICategoryData[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            setIsLoading(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const response = await spHttpClient.get(`${webUrl}/_api/web/lists/getByTitle('ENT_Materials_Master')/items?$top=500`, SPHttpClient.configurations.v1);

                if (!response.ok) {
                    setCategoryData([]);
                    setIsLoading(false);
                    return;
                }

                const data = await response.json();
                const items = data.value || [];

                // Group by category
                const categoryMap = new Map<string, { count: number; totalValue: number }>();

                items.forEach((item: any) => {
                    const category = item.Category || item.Material_Category || 'Uncategorized';
                    const unitPrice = parseFloat(item.Unit_Price || item.Price || '0') || 0;
                    const quantity = parseFloat(item.Quantity || item.Stock_Quantity || '0') || 1;

                    if (categoryMap.has(category)) {
                        const existing = categoryMap.get(category)!;
                        categoryMap.set(category, {
                            count: existing.count + 1,
                            totalValue: existing.totalValue + (unitPrice * quantity),
                        });
                    } else {
                        categoryMap.set(category, { count: 1, totalValue: unitPrice * quantity });
                    }
                });

                const categories: ICategoryData[] = Array.from(categoryMap.entries())
                    .map(([category, data]) => ({
                        category,
                        count: data.count,
                        totalValue: data.totalValue,
                    }))
                    .sort((a, b) => b.count - a.count);

                setCategoryData(categories);
            } catch (error) {
                console.error('Error fetching material category data:', error);
                setCategoryData([]);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [refreshKey, pageContext]);

    const totalMaterials = categoryData.reduce((sum, d) => sum + d.count, 0);
    const totalValue = categoryData.reduce((sum, d) => sum + d.totalValue, 0);
    const maxCount = Math.max(...categoryData.map(d => d.count), 1);

    const colors = [
        theme.palette.blue,
        theme.palette.green,
        theme.palette.orange,
        theme.palette.purple,
        theme.palette.red,
        theme.palette.teal,
        theme.palette.magenta,
        theme.palette.yellow,
    ];

    if (isLoading) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                Loading material category data...
            </div>
        );
    }

    if (categoryData.length === 0) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: theme.palette.neutralSecondary }}>
                No material data available
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            {/* Summary Stats */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '12px' }}>
                <div style={{
                    padding: '12px',
                    backgroundColor: theme.palette.blueLight || `${theme.palette.blue}15`,
                    borderRadius: '8px',
                    borderLeft: `3px solid ${theme.palette.blue}`,
                }}>
                    <div style={{ fontSize: '20px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                        {totalMaterials}
                    </div>
                    <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                        Total Materials
                    </div>
                </div>
                <div style={{
                    padding: '12px',
                    backgroundColor: theme.palette.greenLight || `${theme.palette.green}15`,
                    borderRadius: '8px',
                    borderLeft: `3px solid ${theme.palette.green}`,
                }}>
                    <div style={{ fontSize: '20px', fontWeight: 700, color: theme.palette.neutralPrimary }}>
                        {categoryData.length}
                    </div>
                    <div style={{ fontSize: '12px', color: theme.palette.neutralSecondary }}>
                        Categories
                    </div>
                </div>
            </div>

            {/* Category List with Progress Bars */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {categoryData.slice(0, 8).map((item, index) => {
                    const percent = (item.count / totalMaterials) * 100;
                    const valuePercent = totalValue > 0 ? (item.totalValue / totalValue) * 100 : 0;
                    const color = colors[index % colors.length];

                    return (
                        <div key={item.category} style={{
                            padding: '12px',
                            backgroundColor: theme.palette.white,
                            borderRadius: '8px',
                            border: `1px solid ${theme.palette.neutralLight}`,
                        }}>
                            <div style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center',
                                marginBottom: '8px',
                            }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <div style={{
                                        width: '8px',
                                        height: '8px',
                                        borderRadius: '50%',
                                        backgroundColor: color,
                                    }} />
                                    <span style={{
                                        fontSize: '14px',
                                        fontWeight: 500,
                                        color: theme.palette.neutralPrimary,
                                    }}>
                                        {item.category}
                                    </span>
                                </div>
                                <span style={{ fontSize: '13px', color: theme.palette.neutralSecondary }}>
                                    {item.count} items ({percent.toFixed(1)}%)
                                </span>
                            </div>

                            {/* Progress Bar */}
                            <div style={{
                                height: '6px',
                                backgroundColor: theme.palette.neutralLighter,
                                borderRadius: '3px',
                                overflow: 'hidden',
                                marginBottom: '6px',
                            }}>
                                <div style={{
                                    height: '100%',
                                    width: `${percent}%`,
                                    backgroundColor: color,
                                    borderRadius: '3px',
                                }} />
                            </div>

                            {/* Value */}
                            <div style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                fontSize: '11px',
                                color: theme.palette.neutralTertiary,
                            }}>
                                <span>Inventory Value</span>
                                <span>${item.totalValue.toLocaleString()}</span>
                            </div>
                        </div>
                    );
                })}
            </div>

            {categoryData.length > 8 && (
                <div style={{
                    textAlign: 'center',
                    fontSize: '12px',
                    color: theme.palette.neutralSecondary,
                    paddingTop: '8px',
                }}>
                    +{categoryData.length - 8} more categories
                </div>
            )}
        </div>
    );
};

export default MaterialCategoryBreakdown;
