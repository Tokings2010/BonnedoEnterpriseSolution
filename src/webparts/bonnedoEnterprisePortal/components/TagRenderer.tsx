import * as React from 'react';
import { getTheme } from '@fluentui/react';

export type TagVariant = 'teal' | 'blue' | 'orange' | 'red' | 'purple' | 'green' | 'yellow' | 'gray';

export const TAG_COLORS: Record<TagVariant, { bg: string; color: string }> = {
  teal: { bg: 'rgba(0,120,212,0.08)', color: '#0078D4' },
  blue: { bg: 'rgba(74,144,217,0.08)', color: '#4A90D9' },
  orange: { bg: 'rgba(243,156,18,0.08)', color: '#F39C12' },
  red: { bg: 'rgba(231,76,60,0.08)', color: '#E74C3C' },
  purple: { bg: 'rgba(155,89,182,0.08)', color: '#9B59B6' },
  green: { bg: 'rgba(0,120,212,0.08)', color: '#0078D4' },
  yellow: { bg: 'rgba(241,196,15,0.12)', color: '#B7950B' },
  gray: { bg: '#F3F4F6', color: '#6B7280' },
};

const statusTagMap: Record<string, TagVariant> = {
  active: 'teal',
  approved: 'teal',
  completed: 'blue',
  pending: 'orange',
  rejected: 'red',
  draft: 'gray',
  submitted: 'blue',
  issued: 'teal',
  'in transit': 'orange',
  inactive: 'red',
  blocked: 'red',
  available: 'teal',
  good: 'teal',
  damaged: 'orange',
  reserved: 'orange',
  paid: 'teal',
  low: 'red',
  ok: 'teal',
};

export function getStatusTagVariant(status?: string): TagVariant {
  return statusTagMap[(status || '').toLowerCase().trim()] || 'gray';
}

export interface ITagProps {
  text: string;
  variant?: TagVariant;
  /** Auto-detect variant from known status values */
  autoVariant?: boolean;
  style?: React.CSSProperties;
}

export const Tag: React.FC<ITagProps> = ({ text, variant, autoVariant = true, style }) => {
  const resolved = variant || (autoVariant ? getStatusTagVariant(text) : 'gray');
  const colors = TAG_COLORS[resolved] || TAG_COLORS.gray;
  return (
    <span
      style={{
        display: 'inline-block',
        padding: '3px 10px',
        borderRadius: 20,
        fontSize: 10,
        fontWeight: 600,
        backgroundColor: colors.bg,
        color: colors.color,
        whiteSpace: 'nowrap',
        ...style,
      }}
    >
      {text}
    </span>
  );
};

/** Render a movement type tag (GRN, Transfer Out, Issue, Return, Scrap, Verify) */
export const MovementTag: React.FC<{ type: string }> = ({ type }) => {
  const map: Record<string, TagVariant> = {
    GRN: 'green',
    'Transfer Out': 'blue',
    Issue: 'orange',
    Return: 'purple',
    Scrap: 'red',
    Verify: 'teal',
  };
  return <Tag text={type} variant={map[type] || 'gray'} autoVariant={false} />;
};

/** Progress bar component matching reference design */
export interface IProgressBarProps {
  percent: number;
  color?: string;
  height?: number;
  showLabel?: boolean;
}

export const ProgressBar: React.FC<IProgressBarProps> = ({
  percent,
  color = '#0078D4',
  height = 6,
  showLabel = false,
}) => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
    <span
      style={{
        background: '#E8ECF0',
        borderRadius: 4,
        height,
        width: 90,
        display: 'inline-block',
        verticalAlign: 'middle',
        overflow: 'hidden',
      }}
    >
      <span
        style={{
          display: 'block',
          height,
          borderRadius: 4,
          width: `${Math.min(percent, 100)}%`,
          backgroundColor: color,
          transition: 'width 0.3s ease',
        }}
      />
    </span>
    {showLabel && (
      <span style={{ fontSize: 11, color: '#7F8C9B' }}>{percent.toFixed(0)}%</span>
    )}
  </span>
);

/** Monospace code text */
export const MonoText: React.FC<{ children: React.ReactNode; style?: React.CSSProperties }> = ({
  children,
  style,
}) => (
  <span
    style={{
      fontFamily: "'Cascadia Code', 'Fira Code', Consolas, monospace",
      fontWeight: 600,
      fontSize: 13,
      color: '#2C3E50',
      ...style,
    }}
  >
    {children}
  </span>
);

/** Format a date string to a standard format */
export function formatDate(dateStr?: string): string {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
}

export function formatDateTime(dateStr?: string): string {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr);
    return d.toLocaleString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  } catch {
    return dateStr;
  }
}

/** Format currency */
export function formatCurrency(amount?: number): string {
  if (amount == null) return 'N0';
  if (amount >= 1000000000) return `N${(amount / 1000000000).toFixed(1)}B`;
  if (amount >= 1000000) return `N${(amount / 1000000).toFixed(1)}M`;
  if (amount >= 1000) return `N${(amount / 1000).toFixed(1)}K`;
  return `N${amount.toFixed(0)}`;
}

/** Export data array to CSV file */
export function exportToCsv<T extends Record<string, unknown>>(
  items: T[],
  headers: { key: string; label: string }[],
  filename?: string
): void {
  const esc = (v: unknown): string => {
    const str = v == null ? '' : String(v);
    return `"${str.replace(/"/g, '""')}"`;
  };
  const headerRow = headers.map((h) => esc(h.label)).join(',');
  const dataRows = items.map((item) => headers.map((h) => esc(item[h.key])).join(','));
  const csv = [headerRow, ...dataRows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename || `export-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}
