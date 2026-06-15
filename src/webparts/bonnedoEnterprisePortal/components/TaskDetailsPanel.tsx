import * as React from 'react';
import {
  Panel,
  PanelType,
  Text,
  PrimaryButton,
  DefaultButton,
  Separator,
  Label,
  getTheme,
  mergeStyleSets,
  ScrollablePane,
  ScrollbarVisibility,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import TaskForm from './TaskForm';

export interface ITaskItem {
  ID?: number;
  Title?: string;
  Project_Code?: string;
  Bucket?: string;
  AssignedToEmail?: string;
  DueDate?: string;
  Priority?: string;
  Status?: string;
  Description?: string;
  PlannerTaskId?: string;
  PlannerPlanId?: string;
  AssignedToId?: string;
  [key: string]: any;
}

export interface ITaskDetailsPanelProps {
  isOpen: boolean;
  task: ITaskItem | undefined;
  onDismiss: () => void;
  onRefresh?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext?: any;
}

const TaskDetailsPanel: React.FC<ITaskDetailsPanelProps> = ({
  isOpen,
  task,
  onDismiss,
  onRefresh,
  spHttpClient,
  pageContext,
  webPartContext,
}) => {
  const theme = getTheme();
  const [isEditMode, setIsEditMode] = React.useState(false);

  const handleEdit = (): void => {
    setIsEditMode(true);
  };

  const handleEditFormDismiss = (): void => {
    setIsEditMode(false);
  };

  const handleEditFormSuccess = (): void => {
    setIsEditMode(false);
    if (onRefresh) {
      onRefresh();
    }
    onDismiss();
  };

  const handleMarkComplete = async (): Promise<void> => {
    if (!task?.ID) return;

    try {
      const listName = 'Project_Task_Register';
      const siteUrl = pageContext.web.absoluteUrl;
      const getItemUrl = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items(${task.ID})`;

      // Get etag from response headers
      const getResponse = await spHttpClient.get(getItemUrl, SPHttpClient.configurations.v1);
      const etag = getResponse.headers.get('etag') || getResponse.headers.get('ETag') || '*';

      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items(${task.ID})`;

      await spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': etag,
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({ Status: 'Completed' }),
        }
      );

      // Also update Planner if task has PlannerTaskId and webPartContext
      if (task.PlannerTaskId && webPartContext) {
        try {
          const graphClient = await (webPartContext as any).msGraphClientFactory?.getClient('3');
          if (graphClient) {
            // First, GET the current Planner task to retrieve its @odata.etag (required for Planner updates)
            const currentPlannerTask = await graphClient.api(`planner/tasks/${task.PlannerTaskId}`).get();
            const plannerEtag = currentPlannerTask['@odata.etag'];

            // Update Planner task to completed (percentComplete: 100) with proper If-Match header
            await graphClient
              .api(`planner/tasks/${task.PlannerTaskId}`)
              .header('If-Match', plannerEtag)
              .update({ percentComplete: 100 });
          }
        } catch (plannerError) {
          console.warn('Failed to sync completion to Planner:', plannerError);
          // Continue even if Planner sync fails
        }
      }

      if (onRefresh) {
        onRefresh();
      }
      onDismiss();
    } catch (error) {
      console.error('Error marking task complete:', error);
    }
  };

  const handleDelete = async (): Promise<void> => {
    if (!task?.ID) return;

    if (!window.confirm('Are you sure you want to delete this task?')) {
      return;
    }

    try {
      const listName = 'Project_Task_Register';
      const siteUrl = pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items(${task.ID})`;

      await spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );

      if (onRefresh) {
        onRefresh();
      }
      onDismiss();
    } catch (error) {
      console.error('Error deleting task:', error);
    }
  };

  const classNames = mergeStyleSets({
    panelContent: {
      padding: '20px',
      paddingTop: '32px',
    },
    section: {
      marginBottom: '24px',
    },
    fieldLabel: {
      fontWeight: 600,
      marginBottom: '8px',
      color: theme.palette.neutralPrimary,
    },
    fieldValue: {
      color: theme.palette.neutralSecondary,
      marginBottom: '16px',
      wordBreak: 'break-word',
    },
    statusBadge: {
      display: 'inline-block',
      padding: '6px 12px',
      borderRadius: '4px',
      fontSize: '13px',
      fontWeight: 600,
      marginBottom: '16px',
    },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
      flexWrap: 'wrap',
    },
    actionButton: {
      flex: 1,
      minWidth: '120px',
    },
  });

  if (!task) {
    return null;
  }

   if (isEditMode) {
     return (
       <TaskForm
         isOpen={true}
         onDismiss={handleEditFormDismiss}
         onSubmitSuccess={handleEditFormSuccess}
         spHttpClient={spHttpClient}
         pageContext={pageContext}
         webPartContext={webPartContext}
         initialData={{
           Title: task.Title,
           Project_Code: task.Project_Code,
           AssignedToEmail: task.AssignedToEmail,
           DueDate: task.DueDate,
           Priority: task.Priority,
           Bucket: task.Bucket,
         }}
         existingTask={{
           ID: task.ID,
           Title: task.Title,
           Project_Code: task.Project_Code,
           AssignedToEmail: task.AssignedToEmail,
           DueDate: task.DueDate,
           Priority: task.Priority,
           Bucket: task.Bucket,
           Status: task.Status,
           PlannerTaskId: task.PlannerTaskId,
           PlannerPlanId: task.PlannerPlanId,
           AssignedToId: task.AssignedToId
         }}
       />
     );
   }

  const getStatusColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'completed':
        return theme.palette.green;
      case 'in progress':
        return theme.palette.blue;
      case 'not started':
        return theme.palette.neutralSecondary;
      default:
        return theme.palette.neutralSecondary;
    }
  };

  const getPriorityColor = (priority: string): string => {
    switch (priority?.toLowerCase()) {
      case 'high':
        return theme.palette.red;
      case 'medium':
        return theme.palette.orange;
      case 'low':
        return theme.palette.green;
      default:
        return theme.palette.neutralSecondary;
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Task Details"
      closeButtonAriaLabel="Close"
      isLightDismiss
    >
      <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
        <div className={classNames.panelContent}>
          <div className={classNames.section} style={{ marginTop: '16px' }}>
            <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              {task.Title}
            </Text>
            <div
              className={classNames.statusBadge}
              style={{
                backgroundColor: getStatusColor(task.Status || ''),
                color: theme.palette.white,
              }}
            >
              {task.Status || 'Not Started'}
            </div>
            {task.Priority && (
              <div
                className={classNames.statusBadge}
                style={{
                  backgroundColor: getPriorityColor(task.Priority),
                  color: theme.palette.white,
                  marginLeft: '8px',
                }}
              >
                {task.Priority}
              </div>
            )}
          </div>

          <Separator />

          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Task Information
            </Text>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Project Code</Label>
              <Text className={classNames.fieldValue}>{task.Project_Code || 'N/A'}</Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Bucket</Label>
              <Text className={classNames.fieldValue}>{task.Bucket || 'N/A'}</Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Assigned To</Label>
              <Text className={classNames.fieldValue}>{task.AssignedToEmail || 'N/A'}</Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Due Date</Label>
              <Text className={classNames.fieldValue}>
                {task.DueDate
                  ? new Date(task.DueDate).toLocaleDateString('en-GB', {
                      day: 'numeric',
                      month: 'short',
                      year: 'numeric',
                    })
                  : 'N/A'}
              </Text>
            </div>
          </div>

          <Separator />

          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Actions
            </Text>

            {task.Status === 'Completed' ? (
              <div style={{ 
                padding: '12px', 
                backgroundColor: theme.palette.neutralLighter, 
                borderRadius: '4px',
                textAlign: 'center' 
              }}>
                <Text variant="medium" style={{ color: theme.palette.neutralSecondary }}>
                  ✓ This task is completed and cannot be edited or deleted.
                </Text>
              </div>
            ) : (
              <div className={classNames.buttonGroup}>
                <PrimaryButton
                  text="Edit"
                  onClick={handleEdit}
                  iconProps={{ iconName: 'Edit' }}
                  className={classNames.actionButton}
                />
                <DefaultButton
                  text="Mark Complete"
                  onClick={handleMarkComplete}
                  iconProps={{ iconName: 'CheckMark' }}
                  className={classNames.actionButton}
                />
                <DefaultButton
                  text="Delete"
                  onClick={handleDelete}
                  iconProps={{ iconName: 'Delete' }}
                  className={classNames.actionButton}
                />
              </div>
            )}
          </div>

          {task.Description && (
            <>
              <Separator />
              <div className={classNames.section}>
                <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                  Description
                </Text>
                <Text className={classNames.fieldValue}>{task.Description}</Text>
              </div>
            </>
          )}
        </div>
      </ScrollablePane>
    </Panel>
  );
};

export default TaskDetailsPanel;
