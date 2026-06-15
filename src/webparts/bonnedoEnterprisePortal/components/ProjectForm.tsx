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
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';

export interface IProjectFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmitSuccess?: (projectDetails?: {
    projectId: number;
    projectName: string;
    projectCode: string;
    projectManagerId?: string | number;
  }) => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext?: WebPartContext;
  editMode?: boolean;
  editProject?: {
    ID: number;
    Project_Code: string;
    Project_Name: string;
    Client_Name: string;
    Project_ManagerId: string | number | undefined;
    Contract_Value: number;
    Start_Date: string;
    End_Date: string;
  };
}

interface IProjectFormData {
  Project_Code: string;
  Project_Name: string;
  Client_Name: string;
  Project_ManagerId: string | number | undefined;
  Contract_Value: string;
  Start_Date: string;
  End_Date: string;
}

interface IFormState {
  formData: IProjectFormData;
  isSubmitting: boolean;
  error: string | undefined;
  successMessage: string | undefined;
}

const ProjectForm: React.FC<IProjectFormProps> = ({
  isOpen,
  onDismiss,
  onSubmitSuccess,
  spHttpClient,
  pageContext,
  webPartContext,
  editMode,
  editProject,
}) => {
  const [state, setState] = React.useState<IFormState>({
    formData: {
      Project_Code: '',
      Project_Name: '',
      Client_Name: '',
      Project_ManagerId: undefined,
      Contract_Value: '',
      Start_Date: '',
      End_Date: '',
    },
    isSubmitting: false,
    error: undefined,
    successMessage: undefined,
  });

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  React.useEffect(() => {
    if (isOpen) {
      // If in edit mode, populate form with project data
      if (editMode && editProject) {
        setState({
          formData: {
            Project_Code: editProject.Project_Code || '',
            Project_Name: editProject.Project_Name || '',
            Client_Name: editProject.Client_Name || '',
            Project_ManagerId: editProject.Project_ManagerId,
            Contract_Value: editProject.Contract_Value?.toString() || '',
            Start_Date: editProject.Start_Date ? editProject.Start_Date.split('T')[0] : '',
            End_Date: editProject.End_Date ? editProject.End_Date.split('T')[0] : '',
          },
          isSubmitting: false,
          error: undefined,
          successMessage: undefined,
        });
        return;
      }

      // Auto-generate Project Code for new projects
      // Note: This is now handled dynamically in handleInputChange
      // when the user enters Client Name and Project Name
      const autoGenerateCode = async (): Promise<void> => {
        // No longer needed - code is generated dynamically
        // Keeping this function for potential future use
      };

      // Only auto-generate if the code is currently empty (new form)
      if (!state.formData.Project_Code) {
        autoGenerateCode().catch(console.error);
      }
    }
  }, [isOpen, sharePointService]);

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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleInputChange = (fieldName: keyof IProjectFormData, value: string | number | undefined): void => {
    let updatedFormData = {
      ...state.formData,
      [fieldName]: value,
    };

    // Auto-generate Project Code when Client_Name or Project_Name changes
    if (fieldName === 'Client_Name' || fieldName === 'Project_Name') {
      const currentClientName = fieldName === 'Client_Name' ? (value as string) : state.formData.Client_Name;
      const currentProjectName = fieldName === 'Project_Name' ? (value as string) : state.formData.Project_Name;

      // Only generate code if both fields have values and we're not in edit mode
      if (!editMode && currentClientName && currentProjectName && currentClientName.trim() && currentProjectName.trim()) {
        const year = new Date().getFullYear();
        // Use full Client Name (uppercase, no spaces)
        const clientPart = currentClientName.trim().toUpperCase().replace(/\s+/g, '');
        // Get first 3 letters of Project Name (uppercase, no spaces)
        const projectPart = currentProjectName.trim().toUpperCase().replace(/\s+/g, '').substring(0, 3);

        // Format: PRJ-{ClientPart}-{ProjectPart}-{Year}
        const newCode = `PRJ-${clientPart}-${projectPart}-${year}`;
        updatedFormData = {
          ...updatedFormData,
          Project_Code: newCode,
        };
      }
    }

    setState((prev) => ({
      ...prev,
      formData: updatedFormData,
      error: undefined,
    }));
  };

  const validateForm = (): boolean => {
    if (!state.formData.Project_Code.trim()) {
      setState((prev) => ({ ...prev, error: 'Project Code is required' }));
      return false;
    }

    if (!state.formData.Project_Name.trim()) {
      setState((prev) => ({ ...prev, error: 'Project Name is required' }));
      return false;
    }

    if (!state.formData.Client_Name.trim()) {
      setState((prev) => ({ ...prev, error: 'Client Name is required' }));
      return false;
    }

    if (!state.formData.Project_ManagerId) {
      setState((prev) => ({ ...prev, error: 'Project Manager is required' }));
      return false;
    }

    if (!state.formData.Contract_Value.trim()) {
      setState((prev) => ({ ...prev, error: 'Contract Value is required' }));
      return false;
    }

    if (!state.formData.Start_Date.trim()) {
      setState((prev) => ({ ...prev, error: 'Start Date is required' }));
      return false;
    }

    if (!state.formData.End_Date.trim()) {
      setState((prev) => ({ ...prev, error: 'End Date is required' }));
      return false;
    }

    const startDate = new Date(state.formData.Start_Date);
    const endDate = new Date(state.formData.End_Date);

    if (endDate <= startDate) {
      setState((prev) => ({ ...prev, error: 'End Date must be after Start Date' }));
      return false;
    }

    return true;
  };

  const handleSubmit = React.useCallback(async () => {
    if (!validateForm()) {
      return;
    }

    setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

    try {
      const itemData = {
        Title: state.formData.Project_Code,
        Project_Code: state.formData.Project_Code,
        Project_Name: state.formData.Project_Name,
        Client_Name: state.formData.Client_Name,
        Project_ManagerId: typeof state.formData.Project_ManagerId === 'string'
          ? parseInt(state.formData.Project_ManagerId, 10)
          : state.formData.Project_ManagerId,
        Contract_Value: parseFloat(state.formData.Contract_Value),
        Start_Date: new Date(state.formData.Start_Date).toISOString(),
        End_Date: new Date(state.formData.End_Date).toISOString(),
        Project_Status: 'Active',
      };

      if (editMode && editProject) {
        await sharePointService.updateListItem('ENT_Project_Master', editProject.ID, itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Project updated successfully!',
        }));
      } else {
        const createdItem = await sharePointService.createListItem('ENT_Project_Master', itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Project created successfully!',
          formData: {
            Project_Code: '',
            Project_Name: '',
            Client_Name: '',
            Project_ManagerId: undefined,
            Contract_Value: '',
            Start_Date: '',
            End_Date: '',
          },
        }));

        // Pass project details for provisioning
        if (onSubmitSuccess) {
          onSubmitSuccess({
            projectId: createdItem.ID,
            projectName: state.formData.Project_Name,
            projectCode: state.formData.Project_Code,
            projectManagerId: state.formData.Project_ManagerId,
          });
          return; // Return early to avoid double callback
        }
      }

      // Close panel after success
      setTimeout(() => {
        onDismiss();
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error
        ? error.message
        : (editMode ? 'Failed to update project' : 'Failed to create project');
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [state.formData, sharePointService, onDismiss, onSubmitSuccess]);

  const handleCancel = (): void => {
    setState({
      formData: {
        Project_Code: '',
        Project_Name: '',
        Client_Name: '',
        Project_ManagerId: undefined,
        Contract_Value: '',
        Start_Date: '',
        End_Date: '',
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
      headerText={editMode ? 'Edit Project' : 'New Project'}
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
            {/* Project Code */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Project Code"
                placeholder="Auto-generated"
                value={state.formData.Project_Code}
                disabled={true}
                required
              />
            </div>

            {/* Project Name */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Project Name"
                placeholder="Enter project name"
                value={state.formData.Project_Name}
                onChange={(_, value) => handleInputChange('Project_Name', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Client Name */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Client Name"
                placeholder="Enter client name"
                value={state.formData.Client_Name}
                onChange={(_, value) => handleInputChange('Client_Name', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Project Manager */}
            <div className={classNames.fieldGroup}>
              <PeoplePicker
                titleText="Project Manager"
                selectedUsers={state.formData.Project_ManagerId ? [{ id: typeof state.formData.Project_ManagerId === 'string' ? parseInt(state.formData.Project_ManagerId, 10) : state.formData.Project_ManagerId, Title: '', Email: '', LoginName: '' }] : []}
                onChange={(users) => {
                  if (users && users.length > 0 && users[0].id) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    (handleInputChange as any)('Project_ManagerId', users[0].id);
                  } else {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    (handleInputChange as any)('Project_ManagerId', undefined);
                  }
                }}
                personSelectionLimit={1}
                required={true}
                disabled={isDisabled}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                webPartContext={webPartContext}
              />
            </div>

            {/* Contract Value */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Contract Value (₦)"
                placeholder="Enter contract value"
                type="number"
                value={state.formData.Contract_Value}
                onChange={(_, value) => handleInputChange('Contract_Value', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Start Date */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="Start Date"
                type="date"
                value={state.formData.Start_Date}
                onChange={(_, value) => handleInputChange('Start_Date', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* End Date */}
            <div className={classNames.fieldGroup}>
              <TextField
                label="End Date"
                type="date"
                value={state.formData.End_Date}
                onChange={(_, value) => handleInputChange('End_Date', value)}
                disabled={isDisabled}
                required
              />
            </div>

            {/* Buttons */}
            <div className={classNames.buttonGroup}>
              <PrimaryButton
                text={editMode ? 'Update Project' : 'Create Project'}
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
                <Spinner size={SpinnerSize.medium} label="Creating project..." />
              </div>
            )}
          </Stack>
        </div>
      </div>
    </Panel>
  );
};

export default ProjectForm;
