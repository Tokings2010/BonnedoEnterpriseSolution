import * as React from 'react';
import {
  Text,
  getTheme,
  mergeStyleSets,
  Spinner,
  SpinnerSize,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Icon,
  MessageBar,
  MessageBarType,
  TextField,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';

export interface IActivityProgressTabProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  projectCode: string;
  isMobileView: boolean;
  onRefresh: () => void;
}

interface IWBSActivity {
  ID: number;
  Title: string;
  WBS_ID: string;
  Activity_Name: string;
  Planned_Start: string;
  Planned_Finish: string;
  Actual_Start?: string;
  Actual_Finish?: string;
  Duration_Days: number;
  Percent_Complete: number;
  Status: string;
  Level: number;
  Phase_Link: string;
  Weight: number;
  Predecessors: string;
  Remarks: string;
  Physical_Progress: number;
}

const ActivityProgressTab: React.FC<IActivityProgressTabProps> = ({
  spHttpClient,
  pageContext,
  projectCode,
  isMobileView,
  onRefresh,
}) => {
  const theme = getTheme();
  const [activities, setActivities] = React.useState<IWBSActivity[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [successMsg, setSuccessMsg] = React.useState<string | null>(null);
  const [editingRow, setEditingRow] = React.useState<number | null>(null);
  const [editForm, setEditForm] = React.useState<{
    Status: string;
    Percent_Complete: string;
    Actual_Start: string;
    Actual_Finish: string;
  }>({ Status: '', Percent_Complete: '', Actual_Start: '', Actual_Finish: '' });
  const [saving, setSaving] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      gap: '16px',
    },
    headerRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      flexWrap: 'wrap',
      gap: '12px',
    },
    statsRow: {
      display: 'flex',
      gap: '16px',
      flexWrap: 'wrap',
    },
    statCard: {
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '8px',
      padding: '12px 20px',
      display: 'flex',
      alignItems: 'center',
      gap: '10px',
      border: `1px solid ${theme.palette.neutralLight}`,
    },
    statValue: {
      fontSize: '20px',
      fontWeight: 700,
      color: theme.palette.neutralPrimary,
    },
    statLabel: {
      fontSize: '12px',
      color: theme.palette.neutralSecondary,
    },
    gridContainer: {
      flex: 1,
      overflow: 'hidden',
      display: 'flex',
      flexDirection: 'column',
    },
    editRow: {
      backgroundColor: theme.palette.themeLighterAlt,
      padding: '16px',
      borderRadius: '8px',
      border: `1px solid ${theme.palette.themeLighter}`,
      marginBottom: '8px',
    },
    editRowTitle: {
      fontSize: '13px',
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      marginBottom: '12px',
    },
    editFields: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))',
      gap: '12px',
    },
    editActions: {
      display: 'flex',
      gap: '8px',
      marginTop: '12px',
    },
    summaryBar: {
      display: 'flex',
      gap: '24px',
      padding: '12px 16px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '8px',
      border: `1px solid ${theme.palette.neutralLight}`,
      flexWrap: 'wrap',
    },
    summaryItem: {
      display: 'flex',
      flexDirection: 'column',
      gap: '2px',
    },
    summaryLabel: {
      fontSize: '11px',
      fontWeight: 600,
      color: theme.palette.neutralSecondary,
      textTransform: 'uppercase',
    },
    summaryValue: {
      fontSize: '16px',
      fontWeight: 700,
      color: theme.palette.neutralPrimary,
    },
    emptyState: {
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '12px',
      padding: '40px 24px',
      textAlign: 'center',
      height: '100%',
      minHeight: '320px',
      backgroundColor: theme.palette.neutralLighterAlt,
      border: `1px dashed ${theme.palette.neutralLight}`,
      borderRadius: '4px',
    },
    emptyStateIcon: {
      fontSize: '48px',
      color: theme.palette.neutralTertiary,
    },
    progressBar: {
      width: '100%',
      height: '8px',
      backgroundColor: theme.palette.neutralLight,
      borderRadius: '4px',
      overflow: 'hidden',
    },
    progressFill: {
      height: '100%',
      borderRadius: '4px',
      transition: 'width 0.3s ease',
    },
  });

  const fetchActivities = React.useCallback(async (): Promise<void> => {
    if (!projectCode) return;
    setIsLoading(true);
    setError(null);
    try {
      const escapedProjectCode = projectCode.replace(/'/g, "''");
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$filter=Project_Code eq '${escapedProjectCode}'&$orderby=WBS_ID asc&$top=500`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        // If filter fails, try fetching all and filtering client-side
        const fallbackUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$top=5000`;
        const fallbackResp: SPHttpClientResponse = await spHttpClient.get(
          fallbackUrl,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (fallbackResp.ok) {
          const fallbackData = await fallbackResp.json();
          const filtered = (fallbackData.value || []).filter(
            (item: any) => String(item.Project_Code || '').trim() === projectCode.trim()
          );
          setActivities(filtered.map((item: any) => ({
            ID: item.ID,
            Title: item.Title || '',
            WBS_ID: item.WBS_ID || '',
            Activity_Name: item.Activity_Name || item.Title || '',
            Planned_Start: item.Planned_Start || '',
            Planned_Finish: item.Planned_Finish || '',
            Actual_Start: item.Actual_Start || '',
            Actual_Finish: item.Actual_Finish || '',
            Duration_Days: item.Duration_Days || 0,
            Percent_Complete: item.Percent_Complete ?? 0,
            Status: item.Status || 'Not Started',
            Level: item.Level || 1,
            Phase_Link: item.Phase_Link || '',
            Weight: item.Weight || 0,
            Predecessors: item.Predecessors || '',
            Remarks: item.Remarks || '',
            Physical_Progress: item.Physical_Progress ?? item.Percent_Complete ?? 0,
          })));
          setIsLoading(false);
          return;
        }
        throw new Error(`Failed to fetch activities: ${response.status}`);
      }

      const data = await response.json();
      setActivities((data.value || []).map((item: any) => ({
        ID: item.ID,
        Title: item.Title || '',
        WBS_ID: item.WBS_ID || '',
        Activity_Name: item.Activity_Name || item.Title || '',
        Planned_Start: item.Planned_Start || '',
        Planned_Finish: item.Planned_Finish || '',
        Actual_Start: item.Actual_Start || '',
        Actual_Finish: item.Actual_Finish || '',
        Duration_Days: item.Duration_Days || 0,
        Percent_Complete: item.Percent_Complete ?? 0,
        Status: item.Status || 'Not Started',
        Level: item.Level || 1,
        Phase_Link: item.Phase_Link || '',
        Weight: item.Weight || 0,
        Predecessors: item.Predecessors || '',
        Remarks: item.Remarks || '',
        Physical_Progress: item.Physical_Progress ?? item.Percent_Complete ?? 0,
      })));
    } catch (err) {
      console.error('Error fetching WBS activities:', err);
      setError(err instanceof Error ? err.message : 'Failed to load activities');
    } finally {
      setIsLoading(false);
    }
  }, [spHttpClient, pageContext, projectCode]);

  React.useEffect(() => {
    if (projectCode) {
      fetchActivities().catch(() => undefined);
    }
  }, [projectCode, refreshKey, fetchActivities]);

  const handleRefresh = (): void => {
    setRefreshKey((prev) => prev + 1);
    onRefresh();
  };

  const startEditing = (activity: IWBSActivity): void => {
    setEditingRow(activity.ID);
    setEditForm({
      Status: activity.Status,
      Percent_Complete: String(activity.Percent_Complete ?? 0),
      Actual_Start: activity.Actual_Start || '',
      Actual_Finish: activity.Actual_Finish || '',
    });
    setError(null);
    setSuccessMsg(null);
  };

  const cancelEditing = (): void => {
    setEditingRow(null);
    setEditForm({ Status: '', Percent_Complete: '', Actual_Start: '', Actual_Finish: '' });
    setError(null);
  };

  const saveActivity = async (activityId: number): Promise<void> => {
    setSaving(true);
    setError(null);
    setSuccessMsg(null);

    try {
      const percentVal = Math.min(100, Math.max(0, parseFloat(editForm.Percent_Complete) || 0));

      const body: Record<string, any> = {
        Status: editForm.Status,
        Percent_Complete: percentVal,
        Physical_Progress: percentVal,
      };

      if (editForm.Actual_Start) {
        body.Actual_Start = editForm.Actual_Start;
      }
      if (editForm.Actual_Finish) {
        body.Actual_Finish = editForm.Actual_Finish;
      }

      if (editForm.Status === 'Completed') {
        body.Percent_Complete = 100;
        body.Physical_Progress = 100;
        if (!editForm.Actual_Finish) {
          body.Actual_Finish = new Date().toISOString();
        }
      }

      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Project_Schedule')/items(${activityId})`;
      const response: SPHttpClientResponse = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify(body),
        }
      );

      if (!response.ok) throw new Error(`Failed to update activity: ${response.status}`);

      setSuccessMsg('Activity progress updated successfully');
      cancelEditing();
      handleRefresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to update activity');
    } finally {
      setSaving(false);
    }
  };

  const formatDate = (dateStr: string): string => {
    if (!dateStr) return '\u2014';
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? dateStr : d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  };

  const toInputDate = (dateStr: string): string => {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? '' : d.toISOString().split('T')[0];
  };

  // Compute summary statistics
  const summaryStats = React.useMemo(() => {
    const total = activities.length;
    const completed = activities.filter((a) => a.Status === 'Completed').length;
    const inProgress = activities.filter((a) => a.Status === 'In Progress').length;
    const notStarted = activities.filter((a) => a.Status === 'Not Started').length;
    const delayed = activities.filter((a) => a.Status === 'Delayed').length;
    const avgProgress = total > 0 ? Math.round(activities.reduce((s, a) => s + (a.Percent_Complete || 0), 0) / total) : 0;
    return { total, completed, inProgress, notStarted, delayed, avgProgress };
  }, [activities]);

  const columns: IDataGridColumn[] = [
    {
      key: 'WBS_ID',
      name: 'WBS ID',
      fieldName: 'WBS_ID',
      minWidth: 80,
      isResizable: true,
      onRender: (item: IWBSActivity) => (
        <span style={{ fontFamily: "'Cascadia Code', Consolas, monospace", fontSize: 12, fontWeight: 600 }}>
          {item.WBS_ID}
        </span>
      ),
    },
    {
      key: 'Activity_Name',
      name: 'Activity',
      fieldName: 'Activity_Name',
      minWidth: 180,
      isResizable: true,
    },
    {
      key: 'Status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 110,
      isResizable: true,
      onRender: (item: IWBSActivity) => {
        const colorMap: Record<string, { bg: string; text: string }> = {
          'Not Started': { bg: '#F3F4F6', text: '#6B7280' },
          'In Progress': { bg: '#DBEAFE', text: '#1E40AF' },
          Completed: { bg: '#D1FAE5', text: '#065F46' },
          Delayed: { bg: '#FEE2E2', text: '#991B1B' },
          'On Hold': { bg: '#FEF3C7', text: '#92400E' },
        };
        const c = colorMap[item.Status] || { bg: '#F3F4F6', text: '#6B7280' };
        return (
          <span style={{
            display: 'inline-block',
            padding: '2px 10px',
            borderRadius: '12px',
            fontSize: '11px',
            fontWeight: 600,
            backgroundColor: c.bg,
            color: c.text,
          }}>
            {item.Status}
          </span>
        );
      },
    },
    {
      key: 'Percent_Complete',
      name: '% Complete',
      fieldName: 'Percent_Complete',
      minWidth: 140,
      isResizable: true,
      onRender: (item: IWBSActivity) => {
        const pct = item.Percent_Complete || 0;
        const color = pct >= 100 ? '#107C10' : pct >= 50 ? '#0078D4' : '#D13438';
        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <div style={{ width: '80px', height: '6px', backgroundColor: '#E8ECF0', borderRadius: '3px', overflow: 'hidden' }}>
              <div style={{ width: `${Math.min(100, pct)}%`, height: '100%', backgroundColor: color, borderRadius: '3px', transition: 'width 0.3s ease' }} />
            </div>
            <span style={{ fontWeight: 600, fontSize: '12px', color }}>{pct}%</span>
          </div>
        );
      },
    },
    {
      key: 'Planned_Start',
      name: 'Planned Start',
      fieldName: 'Planned_Start',
      minWidth: 110,
      isResizable: true,
      onRender: (item: IWBSActivity) => <span>{formatDate(item.Planned_Start)}</span>,
    },
    {
      key: 'Planned_Finish',
      name: 'Planned Finish',
      fieldName: 'Planned_Finish',
      minWidth: 110,
      isResizable: true,
      onRender: (item: IWBSActivity) => <span>{formatDate(item.Planned_Finish)}</span>,
    },
    {
      key: 'Actual_Start',
      name: 'Actual Start',
      fieldName: 'Actual_Start',
      minWidth: 110,
      isResizable: true,
      onRender: (item: IWBSActivity) => <span style={{ color: item.Actual_Start ? theme.palette.neutralPrimary : theme.palette.neutralTertiary }}>{formatDate(item.Actual_Start || '')}</span>,
    },
    {
      key: 'Actual_Finish',
      name: 'Actual Finish',
      fieldName: 'Actual_Finish',
      minWidth: 110,
      isResizable: true,
      onRender: (item: IWBSActivity) => <span style={{ color: item.Actual_Finish ? theme.palette.neutralPrimary : theme.palette.neutralTertiary }}>{formatDate(item.Actual_Finish || '')}</span>,
    },
    {
      key: 'Weight',
      name: 'Weight',
      fieldName: 'Weight',
      minWidth: 70,
      isResizable: true,
    },
    {
      key: 'actions',
      name: 'Update',
      fieldName: '',
      minWidth: 80,
      isResizable: false,
      onRender: (item: IWBSActivity) => (
        <IconButton
          iconProps={{ iconName: 'Edit' }}
          title="Update progress"
          ariaLabel="Update progress"
          onClick={(e) => {
            e.stopPropagation();
            startEditing(item);
          }}
          styles={{ root: { height: 28, width: 28 } }}
        />
      ),
    },
  ];

  if (!projectCode) {
    return (
      <div className={classNames.emptyState}>
        <Icon iconName="TimelineProgress" className={classNames.emptyStateIcon} />
        <Text variant="large" style={{ fontWeight: 600, color: theme.palette.neutralPrimary }}>
          Select a project
        </Text>
        <Text variant="medium" style={{ color: theme.palette.neutralSecondary, maxWidth: '420px' }}>
          Choose a project from the Projects tab to view and update WBS activity progress.
        </Text>
      </div>
    );
  }

  return (
    <div className={classNames.root}>
      {/* Header */}
      <div className={classNames.headerRow}>
        <div>
          <Text variant="xxLarge" style={{ fontWeight: 600 }}>
            Activity Progress Update
          </Text>
          <Text variant="medium" style={{ color: theme.palette.neutralSecondary }}>
            Project: <strong>{projectCode}</strong>
          </Text>
        </div>
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          onClick={handleRefresh}
          title="Refresh activities"
          ariaLabel="Refresh activities"
        />
      </div>

      {/* Summary Bar */}
      {activities.length > 0 && (
        <div className={classNames.summaryBar}>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>Activities</span>
            <span className={classNames.summaryValue}>{summaryStats.total}</span>
          </div>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>Avg Progress</span>
            <span className={classNames.summaryValue} style={{ color: summaryStats.avgProgress >= 100 ? '#107C10' : summaryStats.avgProgress >= 50 ? '#0078D4' : '#D13438' }}>
              {summaryStats.avgProgress}%
            </span>
          </div>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>Completed</span>
            <span className={classNames.summaryValue} style={{ color: '#107C10' }}>{summaryStats.completed}</span>
          </div>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>In Progress</span>
            <span className={classNames.summaryValue} style={{ color: '#0078D4' }}>{summaryStats.inProgress}</span>
          </div>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>Delayed</span>
            <span className={classNames.summaryValue} style={{ color: summaryStats.delayed > 0 ? '#D13438' : theme.palette.neutralSecondary }}>{summaryStats.delayed}</span>
          </div>
          <div className={classNames.summaryItem}>
            <span className={classNames.summaryLabel}>Not Started</span>
            <span className={classNames.summaryValue} style={{ color: theme.palette.neutralSecondary }}>{summaryStats.notStarted}</span>
          </div>
        </div>
      )}

      {/* Messages */}
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)} isMultiline={false}>
          {error}
        </MessageBar>
      )}
      {successMsg && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMsg(null)} isMultiline={false}>
          {successMsg}
        </MessageBar>
      )}

      {/* Edit Form */}
      {editingRow !== null && (
        <div className={classNames.editRow}>
          <div className={classNames.editRowTitle}>
            <Icon iconName="Edit" style={{ marginRight: 6 }} />
            Update Progress &mdash; {activities.find((a) => a.ID === editingRow)?.Activity_Name || activities.find((a) => a.ID === editingRow)?.WBS_ID}
          </div>
          <div className={classNames.editFields}>
            <TextField
              label="Status"
              value={editForm.Status}
              onChange={(_, v) => setEditForm({ ...editForm, Status: v || '' })}
              placeholder="e.g. In Progress, Completed"
              disabled={saving}
            />
            <TextField
              label="% Complete"
              type="number"
              min={0}
              max={100}
              step={1}
              value={editForm.Percent_Complete}
              onChange={(_, v) => setEditForm({ ...editForm, Percent_Complete: v || '0' })}
              disabled={saving}
            />
            <TextField
              label="Actual Start"
              type="date"
              value={toInputDate(editForm.Actual_Start)}
              onChange={(_, v) => setEditForm({ ...editForm, Actual_Start: v || '' })}
              disabled={saving}
            />
            <TextField
              label="Actual Finish"
              type="date"
              value={toInputDate(editForm.Actual_Finish)}
              onChange={(_, v) => setEditForm({ ...editForm, Actual_Finish: v || '' })}
              disabled={saving}
            />
          </div>
          <div className={classNames.editActions}>
            <PrimaryButton
              text={saving ? 'Saving...' : 'Save Progress'}
              iconProps={{ iconName: 'Save' }}
              onClick={() => saveActivity(editingRow)}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={cancelEditing}
              disabled={saving}
            />
          </div>
        </div>
      )}

      {/* Data Grid */}
      <div className={classNames.gridContainer}>
        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Loading WBS activities..." />
        ) : activities.length === 0 ? (
          <div className={classNames.emptyState}>
            <Icon iconName="TimelineProgress" className={classNames.emptyStateIcon} />
            <Text variant="large" style={{ fontWeight: 600, color: theme.palette.neutralPrimary }}>
              No WBS activities found
            </Text>
            <Text variant="medium" style={{ color: theme.palette.neutralSecondary, maxWidth: '420px' }}>
              Upload a WBS file or create activities in the Project Schedule tab to track progress.
            </Text>
          </div>
        ) : (
          <EnhancedDataGrid
            key={`progress-${refreshKey}`}
            listName="Project_Schedule"
            columns={columns}
            pageSize={25}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            showExport
          />
        )}
      </div>
    </div>
  );
};

export default ActivityProgressTab;
