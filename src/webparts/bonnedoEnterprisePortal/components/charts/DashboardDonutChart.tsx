import * as React from 'react';
import { getTheme } from '@fluentui/react';

interface IStatusData {
    label: string;
    value: number;
    color: string;
}

interface IDashboardDonutChartProps {
    data: IStatusData[];
    total: number;
}

const DashboardDonutChart: React.FC<IDashboardDonutChartProps> = ({
    data = [],
    total = 0,
}) => {
    const theme = getTheme();

    if (!data || data.length === 0 || total <= 0) {
        return (
            <div style={{ height: 180, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#666' }}>
                No data
            </div>
        );
    }

    const size = 140;
    const strokeWidth = 25;
    const radius = (size - strokeWidth) / 2;
    const circumference = 2 * Math.PI * radius;

    // Calculate offsets for each segment
    let currentOffset = 0;
    const segments = data.map((item) => {
        const percentage = total > 0 ? (item.value / total) : 0;
        const segmentLength = percentage * circumference;
        const offset = currentOffset;
        currentOffset += segmentLength;

        return {
            ...item,
            percentage,
            segmentLength,
            offset,
        };
    });

    // Handle case where there is no data
    if (total === 0) {
        return (
            <div style={{
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                height: '180px'
            }}>
                <div style={{
                    width: size,
                    height: size,
                    borderRadius: '50%',
                    border: `${strokeWidth}px solid ${theme.palette.neutralLight}`,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                }}>
                    <span style={{ fontSize: '14px', color: theme.palette.neutralSecondary }}>No Data</span>
                </div>
            </div>
        );
    }

    return (
        <div style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: '16px',
            overflow: 'hidden',
        }}>
            {/* Donut Chart */}
            <div style={{ position: 'relative', width: size, height: size }}>
                <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
                    {/* Background circle */}
                    <circle
                        cx={size / 2}
                        cy={size / 2}
                        r={radius}
                        fill="none"
                        stroke={theme.palette.neutralLight}
                        strokeWidth={strokeWidth}
                    />

                    {/* Data segments */}
                    {segments.map((segment) => (
                        <circle
                            key={segment.label}
                            cx={size / 2}
                            cy={size / 2}
                            r={radius}
                            fill="none"
                            stroke={segment.color}
                            strokeWidth={strokeWidth}
                            strokeDasharray={`${segment.segmentLength} ${circumference - segment.segmentLength}`}
                            strokeDashoffset={-segment.offset}
                            transform={`rotate(-90 ${size / 2} ${size / 2})`}
                            style={{ transition: 'stroke-dasharray 0.5s ease' }}
                        />
                    ))}
                </svg>

                {/* Center text */}
                <div style={{
                    position: 'absolute',
                    top: '50%',
                    left: '50%',
                    transform: 'translate(-50%, -50%)',
                    textAlign: 'center',
                }}>
                    <span style={{
                        fontSize: '24px',
                        fontWeight: 700,
                        color: theme.palette.neutralPrimary,
                    }}>
                        {total}
                    </span>
                    <br />
                    <span style={{
                        fontSize: '11px',
                        color: theme.palette.neutralSecondary,
                    }}>
                        Total
                    </span>
                </div>
            </div>

            {/* Legend */}
            <div style={{
                display: 'flex',
                justifyContent: 'center',
                gap: '16px',
                flexWrap: 'wrap',
            }}>
                {segments.map((item) => (
                    <div key={item.label} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <div style={{
                            width: '14px',
                            height: '14px',
                            backgroundColor: item.color,
                            borderRadius: '3px'
                        }} />
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontSize: '12px', fontWeight: 500, color: theme.palette.neutralPrimary }}>
                                {item.label}
                            </span>
                            <span style={{ fontSize: '10px', color: theme.palette.neutralSecondary }}>
                                {item.value} ({Math.round(item.percentage * 100)}%)
                            </span>
                        </div>
                    </div>
                ))}
            </div>
        </div>
    );
};

export default DashboardDonutChart;
