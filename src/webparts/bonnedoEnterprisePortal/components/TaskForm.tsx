import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  mergeStyleSets,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';
import { createTaskRegisterService, ITaskRegisterItem } from '../services/TaskRegisterService';

export interface ITaskFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmitSuccess?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext?: WebPartContext;
  initialData?: {
    Title?: string;
    Project_Code?: string;
    AssignedToEmail?: string;
    DueDate?: string;
    Priority?: string;
    Bucket?: string;
  };
  existingTask?: {
    ID?: number;
    Title?: string;
    Project_Code?: string;
    Bucket?: string;
    AssignedToEmail?: string;
    DueDate?: string;
    StartDate?: string;
    Priority?: string;
    Status?: string;
    PlannerTaskId?: string;
    PlannerPlanId?: string;
    AssignedToId?: string;
  };
}

interface ITaskFormData {
  Title: string;
  Project_Code: string;
  Bucket: string;
  AssignedToEmail: string;
  StartDate: string;
  DueDate: string;
  Priority: string;
  Status: string;
}

interface IFormState {
  formData: ITaskFormData;
  isSubmitting: boolean;
  error: string | undefined;
  successMessage: string | undefined;
}



const priorityOptions: IDropdownOption[] = [
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' },
];

const statusOptions: IDropdownOption[] = [
  { key: 'Not Started', text: 'Not Started' },
  { key: 'In Progress', text: 'In Progress' },
  { key: 'Completed', text: 'Completed' },
];

const TaskForm: React.FC<ITaskFormProps> = ({
  isOpen,
  onDismiss,
  onSubmitSuccess,
  spHttpClient,
  pageContext,
  webPartContext,
  initialData,
  existingTask,
}) => {
  const [state, setState] = React.useState<IFormState>({
    formData: {
      Title: '',
      Project_Code: '',
      Bucket: 'Planning',
      AssignedToEmail: '',
      StartDate: '',
      DueDate: '',
      Priority: 'High',
      Status: 'Not Started',
    },
    isSubmitting: false,
    error: undefined,
    successMessage: undefined,
  });

  const [projectOptions, setProjectOptions] = React.useState<IDropdownOption[]>([]);
  const [teamMemberOptions, setTeamMemberOptions] = React.useState<IDropdownOption[]>([]);
  const [bucketOptions, setBucketOptions] = React.useState<IDropdownOption[]>([]);
  const [currentPlanId, setCurrentPlanId] = React.useState<string | null>(null);
  const [currentGroupId, setCurrentGroupId] = React.useState<string | null>(null);
  const [selectedUserId, setSelectedUserId] = React.useState<string>('');
  const taskDataRef = React.useRef<any>({});

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

React.useEffect(() => {
  if (isOpen) {
    const today = new Date();
    const due = new Date();
    due.setDate(due.getDate() + 3);

    // Check if we're editing an existing task
    if (existingTask) {
      setState({
        formData: {
          Title: existingTask.Title || '',
          Project_Code: existingTask.Project_Code || '',
          Bucket: existingTask.Bucket || 'Planning',
          AssignedToEmail: existingTask.AssignedToEmail || '',
          StartDate: existingTask.StartDate ? new Date(existingTask.StartDate).toISOString().split('T')[0] : today.toISOString().split('T')[0],
          DueDate: existingTask.DueDate ? new Date(existingTask.DueDate).toISOString().split('T')[0] : (initialData?.DueDate ? new Date(initialData.DueDate) : due).toISOString().split('T')[0],
          Priority: existingTask.Priority || 'High',
          Status: existingTask.Status || 'Not Started',
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
      });

      // Store IDs for Planner assignment during edit
      if (existingTask.AssignedToId) {
        setSelectedUserId(existingTask.AssignedToId);
        (taskDataRef as any).current = { ...(taskDataRef as any).current, AssignedToId: existingTask.AssignedToId };
      }
    } else {
      setState({
        formData: {
          Title: initialData?.Title || '',
          Project_Code: initialData?.Project_Code || '',
          Bucket: initialData?.Bucket || 'Planning',
          AssignedToEmail: initialData?.AssignedToEmail || '',
          StartDate: today.toISOString().split('T')[0],
          DueDate: (initialData?.DueDate ? new Date(initialData.DueDate) : due).toISOString().split('T')[0],
          Priority: initialData?.Priority || 'High',
          Status: 'Not Started',
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
      });
    }

      // Load projects and bucket choices for dropdown
      loadProjects();
      loadBucketChoices();
    }
  }, [isOpen, initialData, existingTask]);

  const loadProjects = async (): Promise<void> => {
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$select=Project_Code,Project_Name&$orderby=Project_Name&$top=100`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const opts: IDropdownOption[] = data.value.map((p: any) => ({
          key: p.Project_Code,
          text: `${p.Project_Code} - ${p.Project_Name}`,
        }));
        setProjectOptions(opts);
      }
    } catch (e) {
      console.warn('Failed to load projects', e);
    }
  };

  const loadBucketChoices = async (): Promise<void> => {
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('Project_Task_Register')/fields/getByInternalNameOrTitle('Bucket')?$select=Choices`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const data = await resp.json();
        const choices: string[] = data.Choices || [];
        const opts: IDropdownOption[] = choices.map((c: string) => ({ key: c, text: c }));
        setBucketOptions(opts);
      }
    } catch (e) {
      console.warn('Failed to load bucket choices', e);
      // Fallback to static list if fetch fails
      setBucketOptions([
        { key: 'Contract & Mobilization', text: 'Contract & Mobilization' },
        { key: 'Planning', text: 'Planning' },
        { key: 'Execution', text: 'Execution' },
        { key: 'WCC Certification', text: 'WCC Certification' },
        { key: 'Invoicing', text: 'Invoicing' },
        { key: 'Closeout', text: 'Closeout' },
      ]);
    }
  };

  const handleProjectChange = async (option?: IDropdownOption): Promise<void> => {
    if (!option) return;
    const code = option.key as string;
    handleInputChange('Project_Code', code);

    // Fetch project details for Planner ID and Group ID
    if (webPartContext) {
      const taskService = createTaskRegisterService(webPartContext);
      const details = await taskService.getProjectDetails(code);
      if (details) {
        setCurrentPlanId(details.plannerPlanId);
        setCurrentGroupId(details.teamsGroupId);

        // Load team members for restricted assignment
        if (details.teamsGroupId) {
          const members = await taskService.getTeamMembers(details.teamsGroupId);
          const memberOpts: IDropdownOption[] = members.map(m => ({
            key: m.id,
            text: `${m.displayName} (${m.email || 'no email'})`,
            data: { email: m.email, id: m.id }
          }));
          setTeamMemberOptions(memberOpts);
        }
      }
    }
  };

  const classNames = mergeStyleSets({
    formContainer: {
      padding: '20px',
      paddingTop: '28px',
    },
    fieldGroup: {
      marginBottom: '16px',
    },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
      flexWrap: 'wrap',
    },
  });

  const handleInputChange = (fieldName: keyof ITaskFormData, value: string | number | undefined): void => {
    setState((prev) => ({
      ...prev,
      formData: { ...prev.formData, [fieldName]: value },
      error: undefined,
    }));
  };

  const handleDropdownChange = (fieldName: keyof ITaskFormData, option?: IDropdownOption): void => {
    if (option) {
      handleInputChange(fieldName, option.key as string);
    }
  };

  const validateForm = (): boolean => {
    if (!state.formData.Title.trim()) {
      setState((prev) => ({ ...prev, error: 'Task Name is required' }));
      return false;
    }
    if (!state.formData.Project_Code.trim()) {
      setState((prev) => ({ ...prev, error: 'Project Code is required' }));
      return false;
    }
    if (!state.formData.DueDate.trim()) {
      setState((prev) => ({ ...prev, error: 'Due Date is required' }));
      return false;
    }
    return true;
  };

const handleSubmit = React.useCallback(async () => {
  if (!validateForm()) return;

  setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

  try {
    const taskData: any = {
      Title: state.formData.Title,
      Project_Code: state.formData.Project_Code,
      Bucket: state.formData.Bucket,
      AssignedToEmail: state.formData.AssignedToEmail,
      StartDate: state.formData.StartDate ? new Date(state.formData.StartDate).toISOString() : undefined,
      DueDate: state.formData.DueDate ? new Date(state.formData.DueDate).toISOString() : undefined,
      Priority: state.formData.Priority,
      Status: state.formData.Status,
      ...taskDataRef.current
    };

    const taskService = createTaskRegisterService(webPartContext!);

    if (existingTask) {
      // Update mode: ensure we have the ID and any Planner IDs from the existing task
      taskData.ID = existingTask.ID;
      if (existingTask.PlannerTaskId) {
        taskData.PlannerTaskId = existingTask.PlannerTaskId;
      }
      if (existingTask.PlannerPlanId) {
        taskData.PlannerPlanId = existingTask.PlannerPlanId;
      }
      // If we have AssignedToId from the existing task (or from the form via taskDataRef), it's already in taskDataRef
      // But note: in edit mode, we set the AssignedToId in taskDataRef in the useEffect above.

      await taskService.updateTaskInSharePointAndPlanner(taskData);

      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        successMessage: 'Task updated successfully and synced to Planner!',
      }));
    } else {
      // Create mode
      const planId = currentPlanId || await taskService.getProjectPlanId(taskData.Project_Code);
      await taskService.createTaskInSharePointAndPlanner(taskData, planId || undefined);

      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        successMessage: 'Task created successfully and synced to Planner!',
      }));
    }

    if (onSubmitSuccess) {
      onSubmitSuccess();
    }

    setTimeout(() => {
      onDismiss();
    }, 1500);
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Failed to save task';
    setState((prev) => ({
      ...prev,
      isSubmitting: false,
      error: errorMessage,
    }));
  }
}, [state.formData, sharePointService, onDismiss, onSubmitSuccess, webPartContext, existingTask]);

  const handleCancel = (): void => {
    setState({
      formData: {
        Title: '',
        Project_Code: '',
        Bucket: 'Planning',
        AssignedToEmail: '',
        StartDate: '',
        DueDate: '',
        Priority: 'High',
        Status: 'Not Started',
      },
      isSubmitting: false,
      error: undefined,
      successMessage: undefined,
    });
    onDismiss();
  };

  const isDisabled = state.isSubmitting;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleCancel}
      type={PanelType.medium}
      headerText="New Task"
      closeButtonAriaLabel="Close"
      isLightDismiss={!state.isSubmitting}
    >
      <div style={{ padding: '0 20px', height: '100%', overflowY: 'auto' }}>
        <div className={classNames.formContainer}>
          {state.error && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {state.error}
              </MessageBar>
            </div>
          )}

          {state.successMessage && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.success} isMultiline>
                {state.successMessage}
              </MessageBar>
            </div>
          )}

          <Stack tokens={{ childrenGap: 16 }}>
            {/* Task Name */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Task Name"
                placeholder="Enter task name"
                value={state.formData.Title}
                onChange={(_, value) => handleInputChange('Title', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Project Code */}
            <div className={classNames.fieldGroup}>
              <Dropdown
                label="Project Code"
                placeholder="Select project"
                selectedKey={state.formData.Project_Code || undefined}
                onChange={(_, option) => handleProjectChange(option)}
                options={projectOptions}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Bucket */}
            <div className={classNames.fieldGroup}>
              <Dropdown
                label="Bucket"
                selectedKey={state.formData.Bucket}
                onChange={(_, option) => handleDropdownChange('Bucket', option)}
                options={bucketOptions}
                disabled={isDisabled}
              />
            </div>

            {/* Assigned To (Team Members Only) */}
            <div className={classNames.fieldGroup}>
              <Dropdown
                label="Assigned To"
                placeholder="Select team member"
                selectedKey={selectedUserId || undefined}
                onChange={(_, option) => {
                  if (option) {
                    setSelectedUserId(option.key as string);
                    const email = (option as any).data?.email || option.text.split('(')[1]?.replace(')', '') || '';
                    handleInputChange('AssignedToEmail', email);
                    // Store ID for Planner assignment
                    (taskDataRef as any).current = { ...(taskDataRef as any).current, AssignedToId: option.key };
                  }
                }}
                options={teamMemberOptions.length > 0 ? teamMemberOptions : [{ key: '', text: 'Select a project first' }]}
                disabled={isDisabled || teamMemberOptions.length === 0}
              />
            </div>

            {/* Start Date */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Start Date"
                type="date"
                value={state.formData.StartDate}
                onChange={(_, value) => handleInputChange('StartDate', value)}
                disabled={isDisabled}
              />
            </div>

            {/* Due Date */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Due Date"
                type="date"
                value={state.formData.DueDate}
                onChange={(_, value) => handleInputChange('DueDate', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Priority */}
            <div className={classNames.fieldGroup}>
              <Dropdown
                label="Priority"
                selectedKey={state.formData.Priority}
                onChange={(_, option) => handleDropdownChange('Priority', option)}
                options={priorityOptions}
                disabled={isDisabled}
              />
            </div>

            {/* Status */}
            <div className={classNames.fieldGroup}>
              <Dropdown
                label="Status"
                selectedKey={state.formData.Status}
                onChange={(_, option) => handleDropdownChange('Status', option)}
                options={statusOptions}
                disabled={isDisabled}
              />
            </div>

            {/* Buttons */}
            <div className={classNames.buttonGroup}>
              <PrimaryButton
                text="Create Task"
                onClick={handleSubmit}
                disabled={isDisabled}
                iconProps={{ iconName: 'CheckMark' }}
              />
              <DefaultButton
                text="Cancel"
                onClick={handleCancel}
                disabled={isDisabled}
                iconProps={{ iconName: 'Cancel' }}
              />
            </div>

            {state.isSubmitting && (
              <div style={{ textAlign: 'center', marginTop: '16px' }}>
                <Spinner size={SpinnerSize.medium} label="Creating task and syncing to Planner..." />
              </div>
            )}
          </Stack>
        </div>
      </div>
    </Panel>
  );
};

export default TaskForm;
