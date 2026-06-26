import * as React from 'react';
import {
  Text,
  getTheme,
  mergeStyleSets,
  PrimaryButton,
  IconButton,
  Spinner,
  SpinnerSize,
  ProgressIndicator,
  Pivot,
  PivotItem,
  Icon,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';
import ProjectDetailsPanel from './ProjectDetailsPanel';
import ProjectCreatePanel from './ProjectsModule/ProjectCreatePanel';
import ProjectScheduleTab from './ProjectsModule/ProjectScheduleTab';
import ProjectBudgetTab from './ProjectsModule/ProjectBudgetTab';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { createProjectProvisioningService, IProvisioningResult, IProvisioningProgress } from '../services/ProjectProvisioningService';
import { createTaskRegisterService } from '../services/TaskRegisterService';
import TaskForm from './TaskForm';
import TaskDetailsPanel, { ITaskItem } from './TaskDetailsPanel';
import ActivityProgressTab from './ActivityProgressTab';

export interface IProjectsModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  webPartContext?: WebPartContext;
}

export interface IProject {
  ID: number;
  Project_Code: string;
  Project_Name: string;
  Client_Name: string;
  Project_Manager: string | { Title: string; ID: number };
  Contract_Value: number;
  Project_Status: string;
  Start_Date: string;
  End_Date: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

const ProjectsModule: React.FC<IProjectsModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  webPartContext,
}) => {
  const theme = getTheme();
  const [selectedProject, setSelectedProject] = React.useState<IProject | undefined>(undefined);
  const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
  const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);
  const [isProvisioning, setIsProvisioning] = React.useState(false);
  const [provisioningProgress, setProvisioningProgress] = React.useState<IProvisioningProgress | null>(null);
  const [isTaskFormOpen, setIsTaskFormOpen] = React.useState(false);
  const [taskInitialData, setTaskInitialData] = React.useState<any>(null);
  const [selectedTab, setSelectedTab] = React.useState<'projects' | 'tasks' | 'schedule' | 'budget' | 'activity-progress'>('projects');
  const [selectedTask, setSelectedTask] = React.useState<ITaskItem | undefined>(undefined);
  const [isTaskDetailsPanelOpen, setIsTaskDetailsPanelOpen] = React.useState(false);

  // Handle window resize for responsive behavior
  React.useEffect(() => {
    const handleResize = (): void => {
      setIsMobileView(window.innerWidth < 768);
    };

    window.addEventListener('resize', handleResize);
    return (): void => {
      window.removeEventListener('resize', handleResize);
    };
  }, []);

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: theme.palette.white,
    },
    header: {
      marginBottom: '20px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      flexWrap: 'wrap',
      gap: '12px',
    },
    headerTitle: {
      flex: 1,
      minWidth: '200px',
    },
    headerActions: {
      display: 'flex',
      gap: '8px',
      flexWrap: 'wrap',
    },
    pivot: {
      flex: 1,
      display: 'flex',
      flexDirection: 'column',
      overflow: 'auto',
    },

    gridContainer: {
      flex: 1,
      overflow: 'hidden',
      display: 'flex',
      flexDirection: 'column',
    },
    cardContainer: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
      gap: '16px',
      overflow: 'auto',
      padding: '8px',
    },
    card: {
      padding: '16px',
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: '4px',
      backgroundColor: theme.palette.white,
      boxShadow: theme.effects.elevation4,
      cursor: 'pointer',
      transition: 'all 0.2s ease',
      selectors: {
        '&:hover': {
          boxShadow: theme.effects.elevation8,
          transform: 'translateY(-2px)',
        },
      },
    },
    cardField: {
      marginBottom: '12px',
      paddingBottom: '12px',
      borderBottom: `1px solid ${theme.palette.neutralLighter}`,
      selectors: {
        '&:last-child': {
          borderBottom: 'none',
          marginBottom: 0,
          paddingBottom: 0,
        },
      },
    },
    cardLabel: {
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
      fontSize: '12px',
      marginBottom: '4px',
    },
    cardValue: {
      color: theme.palette.neutralSecondary,
      fontSize: '14px',
    },
    emptyState: {
      display: 'flex',
      flexDirection: 'column' as const,
      alignItems: 'center',
      justifyContent: 'center',
      gap: '12px',
      padding: '40px 24px',
      textAlign: 'center' as const,
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
    statusBadge: {
      display: 'inline-block',
      padding: '4px 8px',
      borderRadius: '3px',
      fontSize: '12px',
      fontWeight: 600,
    },
    provisioningOverlay: {
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0, 0, 0, 0.6)',
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      zIndex: 9999,
    },
    provisioningCard: {
      backgroundColor: theme.palette.white,
      padding: '40px',
      borderRadius: '12px',
      boxShadow: theme.effects.elevation16,
      minWidth: '420px',
      maxWidth: '520px',
      textAlign: 'center',
      border: `1px solid ${theme.palette.neutralLight}`,
    },
    provisioningHeader: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      marginBottom: '24px',
      gap: '12px',
    },
    provisioningIcon: {
      fontSize: '32px',
      color: theme.palette.themePrimary,
    },
    provisioningTitle: {
      fontWeight: 600,
      color: theme.palette.neutralPrimary,
    },
    provisioningMessage: {
      marginBottom: '28px',
      color: theme.palette.neutralSecondary,
      fontSize: '14px',
      lineHeight: '1.5',
    },
    progressContainer: {
      marginBottom: '20px',
    },
    spinnerContainer: {
      marginTop: '16px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '8px',
    },
  });

  // Define columns for Projects
  const projectColumns: IDataGridColumn[] = [
    {
      key: 'Project_Code',
      name: 'Project Code',
      fieldName: 'Project_Code',
      minWidth: 120,
      isResizable: true,
    },
    {
      key: 'Project_Name',
      name: 'Project Name',
      fieldName: 'Project_Name',
      minWidth: 180,
      isResizable: true,
    },
    {
      key: 'Client_Name',
      name: 'Client Name',
      fieldName: 'Client_Name',
      minWidth: 150,
      isResizable: true,
    },
    {
      key: 'Project_Manager',
      name: 'Project Manager',
      fieldName: 'Project_Manager',
      minWidth: 150,
      isResizable: true,
      onRender: (item: IProject) => {
        const manager = item.Project_Manager;
        const managerName = typeof manager === 'object' ? manager?.Title : manager;
        return <span>{managerName || item.Project_ManagerId || 'N/A'}</span>;
      }
    },
    {
      key: 'Contract_Value',
      name: 'Contract Value',
      fieldName: 'Contract_Value',
      minWidth: 130,
      isResizable: true,
    },
    {
      key: 'Project_Status',
      name: 'Status',
      fieldName: 'Project_Status',
      minWidth: 100,
      isResizable: true,
    },
    {
      key: 'Start_Date',
      name: 'Start Date',
      fieldName: 'Start_Date',
      minWidth: 120,
      isResizable: true,
      onRender: (item: IProject) => <span>{formatDateStandard(item.Start_Date)}</span>,
    },
    {
      key: 'End_Date',
      name: 'End Date',
      fieldName: 'End_Date',
      minWidth: 120,
      isResizable: true,
      onRender: (item: IProject) => <span>{formatDateStandard(item.End_Date)}</span>,
    },
  ];

  const handleRowSelected = (project: IProject): void => {
    setSelectedProject(project);
    setIsDetailsPanelOpen(true);
  };

  const handleRefresh = (): void => {
    setRefreshKey((prev) => prev + 1);
  };

  const handleNewProject = (): void => {
    setIsFormPanelOpen(true);
  };

  const handleFormSubmit = async (projectDetails?: {
    projectId: number;
    projectName: string;
    projectCode: string;
    projectManagerId?: string | number;
    projectManagerEmail?: string;
  }): Promise<void> => {
    setIsFormPanelOpen(false);

    // If project details are provided and we have webPartContext, trigger provisioning
    if (projectDetails && webPartContext) {
      setIsProvisioning(true);
      setProvisioningProgress({ step: 'start', message: 'Starting provisioning...', progress: 0 });

      try {
        console.log('[ProjectsModule] Starting project workspace provisioning for:', projectDetails);

        // Create the provisioning service instance
        const provisioningService = createProjectProvisioningService(webPartContext);

        // Call the provisioning service with progress callback
        const result: IProvisioningResult = await provisioningService.provisionProjectWorkspace(
          {
            projectId: projectDetails.projectId,
            projectName: projectDetails.projectName,
            projectCode: projectDetails.projectCode,
            projectManagerEmail: projectDetails.projectManagerEmail || '',
          },
          (progress: IProvisioningProgress) => {
            setProvisioningProgress(progress);
          }
        );

        if (result.success) {
          console.log('[ProjectsModule] Project workspace provisioned successfully:', result);
          setProvisioningProgress({ step: 'complete', message: 'Provisioning completed successfully!', progress: 100 });
        } else {
          console.error('[ProjectsModule] Project workspace provisioning failed:', result.error);
          setProvisioningProgress({ step: 'error', message: `Provisioning failed: ${result.error}`, progress: 0 });
        }
      } catch (error) {
        console.error('[ProjectsModule] Error during project provisioning:', error);
        setProvisioningProgress({ step: 'error', message: 'Provisioning failed with an error', progress: 0 });
      } finally {
        // Keep the progress visible for a moment before hiding
        setTimeout(() => {
          setIsProvisioning(false);
          setProvisioningProgress(null);
        }, 3000);
      }
    }

    handleRefresh();
  };

  const handleTaskRowSelected = (task: any): void => {
    setSelectedTask(task);
    setIsTaskDetailsPanelOpen(true);
  };

  const handleCreateSampleTask = (): void => {
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 3);
    setTaskInitialData({
      Title: 'Site Survey',
      Project_Code: 'BONNY-PIPE-UPG',
      AssignedToEmail: 'you@contoso.com',
      DueDate: dueDate.toISOString(),
      Priority: 'High',
      Bucket: 'Planning',
    });
    setIsTaskFormOpen(true);
  };

  const formatDateStandard = (dateStr: string): string => {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    const day = date.getDate();
    const month = date.toLocaleString('en-US', { month: 'short' });
    const year = date.getFullYear();
    return `${day} ${month} ${year}`;
  };

  const renderProjectSelectionState = (iconName: string, title: string, message: string): React.ReactNode => (
    <div className={classNames.emptyState}>
      <Icon iconName={iconName} className={classNames.emptyStateIcon} />
      <Text variant="large" block style={{ fontWeight: 600, color: theme.palette.neutralPrimary }}>
        {title}
      </Text>
      <Text variant="medium" style={{ color: theme.palette.neutralSecondary, maxWidth: '420px' }}>
        {message}
      </Text>
    </div>
  );

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item) => {
          if (item?.props.itemKey) {
            setSelectedTab(item.props.itemKey as 'projects' | 'tasks' | 'schedule' | 'budget' | 'activity-progress');
          }
        }}
        className={classNames.pivot || ''}
      >
        {/* Projects Tab */}
        <PivotItem headerText="Projects" itemKey="projects" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            {/* Header Section */}
            <div className={classNames.header}>
              <div className={classNames.headerTitle}>
                <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                  Project Management
                </Text>
                <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                  Manage and track all your projects
                </Text>
              </div>
              <div className={classNames.headerActions}>
                <PrimaryButton
                  text="+ New Project"
                  onClick={handleNewProject}
                  iconProps={{ iconName: 'Add' }}
                />
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={handleRefresh}
                  title="Refresh projects"
                  ariaLabel="Refresh projects"
                />
              </div>
            </div>

            {/* Grid/Card Container */}
            <div className={classNames.gridContainer}>
          <EnhancedDataGrid
                  key={`projects-${refreshKey}`}
                  listName="ENT_Project_Master"
                  columns={projectColumns}
                  pageSize={20}
                  spHttpClient={spHttpClient}
                  pageContext={pageContext}
                  expandQuery="Project_Manager"
                  onRowSelected={handleRowSelected}
                  showExport
                />
            </div>
          </div>
        </PivotItem>

        {/* Project Tasks Tab */}
        <PivotItem headerText="Project Tasks" itemKey="tasks" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            {/* Task Header */}
            <div className={classNames.header}>
              <div className={classNames.headerTitle}>
                <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                  Project Task Register
                </Text>
                <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                  Manage tasks with bidirectional Planner sync
                </Text>
              </div>
              <div className={classNames.headerActions}>
                <PrimaryButton
                  text="+ New Task"
                  onClick={handleCreateSampleTask}
                  iconProps={{ iconName: 'Add' }}
                />
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={handleRefresh}
                  title="Refresh tasks"
                  ariaLabel="Refresh tasks"
                />
              </div>
            </div>

            {/* Task Grid using same DataGrid pattern */}
            <div className={classNames.gridContainer}>
              <EnhancedDataGrid
                key={`tasks-${refreshKey}`}
                listName="Project_Task_Register"
                columns={[
                  { key: 'Title', name: 'Task Name', fieldName: 'Title', minWidth: 180, isResizable: true },
                  { key: 'Project_Code', name: 'Project Code', fieldName: 'Project_Code', minWidth: 120, isResizable: true },
                  { key: 'Bucket', name: 'Bucket', fieldName: 'Bucket', minWidth: 120, isResizable: true },
                  { key: 'AssignedToEmail', name: 'Assigned To', fieldName: 'AssignedToEmail', minWidth: 150, isResizable: true },
                   { key: 'DueDate', name: 'Due Date', fieldName: 'DueDate', minWidth: 120, isResizable: true, onRender: (item: any) => <span>{formatDateStandard(item.DueDate)}</span> },
                  { key: 'Priority', name: 'Priority', fieldName: 'Priority', minWidth: 100, isResizable: true },
                  { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 110, isResizable: true },
                ]}
                pageSize={20}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                 onRowSelected={handleTaskRowSelected}
                 showExport
              />
            </div>
          </div>
        </PivotItem>
        {/* Project Schedule Tab */}
        <PivotItem headerText="Project Schedule" itemKey="schedule" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            <div className={classNames.header}>
              <div className={classNames.headerTitle}>
                <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                  Project Schedule
                </Text>
                <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                  Track WBS activities, progress, and schedule variance
                </Text>
              </div>
              <div className={classNames.headerActions}>
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={handleRefresh}
                  title="Refresh schedule"
                  ariaLabel="Refresh schedule"
                />
              </div>
            </div>

            <div className={classNames.gridContainer}>
              {selectedProject?.Project_Code ? (
                <ProjectScheduleTab
                  projectCode={selectedProject.Project_Code}
                  spHttpClient={spHttpClient}
                  pageContext={pageContext}
                  isMobileView={isMobileView}
                  onRefresh={handleRefresh}
                />
              ) : (
                renderProjectSelectionState('Timeline', 'Select a project', 'Choose a project from the Projects tab to view its schedule and WBS activities.')
              )}
            </div>
          </div>
        </PivotItem>

        {/* Project Budget Tab */}
        <PivotItem headerText="Project Budget" itemKey="budget" style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            <div className={classNames.header}>
              <div className={classNames.headerTitle}>
                <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                  Project Budget
                </Text>
                <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                  Monitor CBS budget, actual spend, variance, and utilization
                </Text>
              </div>
              <div className={classNames.headerActions}>
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={handleRefresh}
                  title="Refresh budget"
                  ariaLabel="Refresh budget"
                />
              </div>
            </div>

            <div className={classNames.gridContainer}>
              {selectedProject?.Project_Code ? (
                <ProjectBudgetTab
                  projectCode={selectedProject.Project_Code}
                  spHttpClient={spHttpClient}
                  pageContext={pageContext}
                  isMobileView={isMobileView}
                  onRefresh={handleRefresh}
                />
              ) : (
                renderProjectSelectionState('Money', 'Select a project', 'Choose a project from the Projects tab to view its budget and CBS data.')
              )}
            </div>
          </div>
        </PivotItem>

        {/* Activity Progress Tab */}
        <PivotItem
          headerText="Progress Update"
          itemKey="activity-progress"
          itemIcon="ProgressRingDots"
          style={{ flex: 1, overflow: 'hidden' }}
        >
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            <div className={classNames.gridContainer}>
              {selectedProject?.Project_Code ? (
                <ActivityProgressTab
                  spHttpClient={spHttpClient}
                  pageContext={pageContext}
                  projectCode={selectedProject.Project_Code}
                  isMobileView={isMobileView}
                  onRefresh={handleRefresh}
                />
              ) : (
                renderProjectSelectionState('ProgressRingDots', 'Select a project', 'Choose a project from the Projects tab to update WBS activity progress.')
              )}
            </div>
          </div>
        </PivotItem>
      </Pivot>

      {/* Project Details Panel */}
      <ProjectDetailsPanel
        isOpen={isDetailsPanelOpen}
        project={selectedProject}
        onDismiss={() => setIsDetailsPanelOpen(false)}
        onRefresh={handleRefresh}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
      />

      {/* Project Create Panel */}
      <ProjectCreatePanel
        isOpen={isFormPanelOpen}
        onDismiss={() => setIsFormPanelOpen(false)}
        onProjectCreated={handleFormSubmit}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
      />

      {/* Task Register Form Panel */}
      <TaskForm
        isOpen={isTaskFormOpen}
        onDismiss={() => {
          setIsTaskFormOpen(false);
          setTaskInitialData(null);
        }}
        onSubmitSuccess={handleRefresh}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
        initialData={taskInitialData}
      />

      {/* Task Details Panel */}
      <TaskDetailsPanel
        isOpen={isTaskDetailsPanelOpen}
        task={selectedTask}
        onDismiss={() => setIsTaskDetailsPanelOpen(false)}
        onRefresh={handleRefresh}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
      />

      {/* Provisioning Progress Overlay */}
      {isProvisioning && provisioningProgress && (
        <div className={classNames.provisioningOverlay}>
          <div className={classNames.provisioningCard}>
            <div className={classNames.provisioningHeader}>
              <span className={classNames.provisioningIcon}>⚙️</span>
              <Text variant="xLarge" className={classNames.provisioningTitle}>
                Provisioning Project Workspace
              </Text>
            </div>
            <Text variant="medium" className={classNames.provisioningMessage}>
              {provisioningProgress.message}
            </Text>
            <div className={classNames.progressContainer}>
              <ProgressIndicator
                percentComplete={provisioningProgress.progress / 100}
                description={`${provisioningProgress.progress}% complete`}
                barHeight={8}
              />
            </div>
            {provisioningProgress.progress < 100 && provisioningProgress.step !== 'error' && (
              <div className={classNames.spinnerContainer}>
                <Spinner size={SpinnerSize.small} />
                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                  This may take a few moments...
                </Text>
              </div>
            )}
            {provisioningProgress.step === 'error' && (
              <Text variant="small" style={{ color: theme.palette.red, marginTop: '12px' }}>
                An error occurred. Please check the console for details.
              </Text>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default ProjectsModule;
