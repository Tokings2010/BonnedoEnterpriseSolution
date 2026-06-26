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
import { IProject } from './ProjectsModule';
import ProjectCreatePanel from './ProjectsModule/ProjectCreatePanel';
import ApprovalTrackerPanel from './ApprovalTrackerPanel';

export interface IProjectDetailsPanelProps {
  isOpen: boolean;
  project: IProject | undefined;
  onDismiss: () => void;
  onRefresh?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

const ProjectDetailsPanel: React.FC<IProjectDetailsPanelProps> = ({
  isOpen,
  project,
  onDismiss,
  onRefresh,
  spHttpClient,
  pageContext,
}) => {
  const theme = getTheme();
  const [isEditMode, setIsEditMode] = React.useState(false);
  const [isApprovalTrackerOpen, setIsApprovalTrackerOpen] = React.useState(false);

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

  const openApprovalTracker = (): void => {
    setIsApprovalTrackerOpen(true);
  };

  const closeApprovalTracker = (): void => {
    setIsApprovalTrackerOpen(false);
  };

  const handleApprovalRefresh = (): void => {
    if (onRefresh) {
      onRefresh();
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

  if (!project) {
    return null;
  }

  // Show edit form when in edit mode
  if (isEditMode) {
    return (
      <ProjectCreatePanel
        isOpen={true}
        onDismiss={handleEditFormDismiss}
        onProjectUpdated={handleEditFormSuccess}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        editMode={true}
        editProject={{
          ID: project.ID,
          Project_Code: typeof project.Project_Code === 'object' ? '' : project.Project_Code,
          Project_Name: typeof project.Project_Name === 'object' ? '' : project.Project_Name,
          Client_Name: typeof project.Client_Name === 'object' ? '' : project.Client_Name,
          Project_ManagerId: project.Project_ManagerId || (typeof project.Project_Manager === 'object' ? project.Project_Manager?.ID : undefined),
          Project_Manager: project.Project_Manager,
          Contract_Value: typeof project.Contract_Value === 'number' ? project.Contract_Value : parseFloat(String(project.Contract_Value || '0')),
          Start_Date: typeof project.Start_Date === 'string' ? project.Start_Date : '',
          End_Date: typeof project.End_Date === 'string' ? project.End_Date : '',
          Project_Status: project.Project_Status,
        }}
      />
    );
  }

  const getStatusColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'active':
      case 'ongoing':
        return theme.palette.green;
      case 'completed':
        return theme.palette.blue;
      case 'on hold':
      case 'paused':
        return theme.palette.orange;
      case 'cancelled':
        return theme.palette.red;
      default:
        return theme.palette.neutralSecondary;
    }
  };

  const handleOpenTeam = (): void => {
    // Open Teams channel for the project team
    const teamUrl = `https://teams.microsoft.com/l/channel/General/Project-${encodeURIComponent(project.Project_Code || '')}?context=${encodeURIComponent(JSON.stringify({ projectId: project.ID, projectName: project.Project_Name }))}`;
    window.open(teamUrl, '_blank');
  };

  const handleOpenPlanner = (): void => {
    // Open Microsoft Planner for the project
    const plannerUrl = `https://tasks.office.com/${pageContext.aadInfo?.domainId || 'earth'}/Home/Task/${encodeURIComponent(project.Project_Code || '')}?projectId=${project.ID}`;
    window.open(plannerUrl, '_blank');
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Project Details"
      closeButtonAriaLabel="Close"
      isLightDismiss
    >
      <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
        <div className={classNames.panelContent}>
          {/* Project Header */}
          <div className={classNames.section} style={{ marginTop: '16px' }}>
            <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              {project.Project_Name}
            </Text>
            <div
              className={classNames.statusBadge}
              style={{
                backgroundColor: getStatusColor(project.Project_Status),
                color: theme.palette.white,
              }}
            >
              {project.Project_Status}
            </div>
          </div>

          <Separator />

          {/* Project Information */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Project Information
            </Text>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Project Code</Label>
              <Text className={classNames.fieldValue}>{project.Project_Code}</Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Client Name</Label>
              <Text className={classNames.fieldValue}>{project.Client_Name || 'N/A'}</Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Project Manager</Label>
              <Text className={classNames.fieldValue}>
                {typeof project.Project_Manager === 'object' ? project.Project_Manager?.Title : (project.Project_Manager || project.Project_ManagerId || 'N/A')}
              </Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Contract Value</Label>
              <Text className={classNames.fieldValue}>
                ₦{(project.Contract_Value || 0).toLocaleString('en-NG', {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}
              </Text>
            </div>
          </div>

          <Separator />

          {/* Timeline Information */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Timeline
            </Text>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Start Date</Label>
              <Text className={classNames.fieldValue}>
                {project.Start_Date
                  ? new Date(project.Start_Date).toLocaleDateString('en-NG', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric',
                  })
                  : 'N/A'}
              </Text>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>End Date</Label>
              <Text className={classNames.fieldValue}>
                {project.End_Date
                  ? new Date(project.End_Date).toLocaleDateString('en-NG', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric',
                  })
                  : 'N/A'}
              </Text>
            </div>

            {project.Start_Date && project.End_Date && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Duration</Label>
                <Text className={classNames.fieldValue}>
                  {Math.ceil(
                    (new Date(project.End_Date).getTime() -
                      new Date(project.Start_Date).getTime()) /
                    (1000 * 60 * 60 * 24)
                  )}{' '}
                  days
                </Text>
              </div>
            )}
          </div>

          <Separator />

          {/* Budget and Schedule */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Budget & Schedule
            </Text>
            <Text className={classNames.fieldValue}>
              Budget (CBS) and schedule (WBS) files uploaded from the create/edit panel are available in the Project Budget and Project Schedule tabs.
            </Text>
          </div>

          <Separator />

          {/* Action Buttons */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Actions
            </Text>

            <div className={classNames.buttonGroup}>
              <PrimaryButton
                text="Edit"
                onClick={handleEdit}
                iconProps={{ iconName: 'Edit' }}
                className={classNames.actionButton}
              />
              <DefaultButton
                text="Team"
                onClick={handleOpenTeam}
                iconProps={{ iconName: 'People' }}
                className={classNames.actionButton}
              />
              <DefaultButton
                text="Planner"
                onClick={handleOpenPlanner}
                iconProps={{ iconName: 'Calendar' }}
                className={classNames.actionButton}
              />
            </div>

            {/* Track Approval Button */}
            <div style={{ marginTop: '16px', paddingTop: '16px', borderTop: `1px solid ${theme.palette.neutralLight}` }}>
              <DefaultButton
                text="Track Approval"
                iconProps={{ iconName: 'CompletedSolid' }}
                onClick={openApprovalTracker}
                style={{ width: '100%' }}
              />
            </div>
          </div>

          {/* Additional Information */}
          {project.Description && (
            <>
              <Separator />
              <div className={classNames.section}>
                <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                  Description
                </Text>
                <Text className={classNames.fieldValue}>{project.Description}</Text>
              </div>
            </>
          )}
        </div>
      </ScrollablePane>

      {/* Approval Tracker Panel */}
      {project && isApprovalTrackerOpen && (
        <ApprovalTrackerPanel
          isOpen={isApprovalTrackerOpen}
          recordId={project.ID}
          listName="ENT_Project_Master"
          spHttpClient={spHttpClient}
          pageContext={pageContext}
          onDismiss={closeApprovalTracker}
          onRefresh={handleApprovalRefresh}
        />
      )}
    </Panel>
  );
};

export default ProjectDetailsPanel;
