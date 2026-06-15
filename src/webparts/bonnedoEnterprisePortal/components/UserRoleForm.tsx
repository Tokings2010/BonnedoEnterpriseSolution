import * as React from 'react';
import {
    Panel,
    PanelType,
    Stack,
    TextField,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    DefaultButton,
    MessageBar,
    MessageBarType,
    Text,
    mergeStyleSets,
    getTheme,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';
import PeoplePicker from './PeoplePicker';

export interface IUserRoleFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
    editMode: boolean;
    editUserRole?: {
        ID: number;
        User: { Title: string; ID: number };
        Role: string;
        Department: string;
        Permissions: string;
    };
}

interface IUserRoleFormData {
    UserId: number | string | undefined;
    Role: string;
    Department: string;
    Permissions: string;
}

interface IFormState {
    formData: IUserRoleFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
}

// Helper type for handling input changes
type FormFieldName = 'UserId' | 'Role' | 'Department' | 'Permissions';

// Role options - matches SYS_User_Roles Choice column
const roleOptions: IDropdownOption[] = [
    { key: 'Admin', text: 'Admin' },
    { key: 'Editor', text: 'Editor' },
    { key: 'Approver', text: 'Approver' },
    { key: 'Viewer', text: 'Viewer' },
];

// Department options - matches SYS_User_Roles Choice column
// Add your actual department choices here
const departmentOptions: IDropdownOption[] = [
    { key: 'IT', text: 'IT' },
    { key: 'Finance', text: 'Finance' },
    { key: 'Procurement', text: 'Procurement' },
    { key: 'Operations', text: 'Operations' },
    { key: 'Human Resources', text: 'Human Resources' },
    { key: 'Marketing', text: 'Marketing' },
    { key: 'Sales', text: 'Sales' },
    { key: 'Engineering', text: 'Engineering' },
];

const UserRoleForm: React.FC<IUserRoleFormProps> = ({
    isOpen,
    onDismiss,
    onSubmitSuccess,
    spHttpClient,
    pageContext,
    webPartContext,
    editMode,
    editUserRole,
}) => {
    const theme = getTheme();
    const [sharePointService] = React.useState(() => new SharePointService(spHttpClient, pageContext));

    const [state, setState] = React.useState<IFormState>({
        formData: {
            UserId: undefined,
            Role: '',
            Department: '',
            Permissions: '',
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
    });

    // Initialize form data when opening in edit mode
    React.useEffect(() => {
        if (isOpen && editMode && editUserRole) {
            setState({
                formData: {
                    UserId: editUserRole.User?.ID,
                    Role: editUserRole.Role || '',
                    Department: editUserRole.Department || '',
                    Permissions: editUserRole.Permissions || '',
                },
                isSubmitting: false,
                error: undefined,
                successMessage: undefined,
            });
        } else if (isOpen && !editMode) {
            // Reset form for new entry
            setState({
                formData: {
                    UserId: undefined,
                    Role: '',
                    Department: '',
                    Permissions: '',
                },
                isSubmitting: false,
                error: undefined,
                successMessage: undefined,
            });
        }
    }, [isOpen, editMode, editUserRole]);

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
        userInfo: {
            padding: '12px 16px',
            backgroundColor: theme.palette.neutralLighterAlt,
            borderRadius: '4px',
            marginBottom: '16px',
        },
    });

    const handleInputChange = (fieldName: keyof IUserRoleFormData, value: string | number | undefined): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [fieldName]: value,
            },
            error: undefined,
        }));
    };

    const validateForm = (): boolean => {
        if (!editMode && !state.formData.UserId) {
            setState((prev) => ({ ...prev, error: 'Please select a user' }));
            return false;
        }
        if (!state.formData.Role) {
            setState((prev) => ({ ...prev, error: 'Role is required' }));
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
                Role: state.formData.Role,
                Department: state.formData.Department,
                Permissions: state.formData.Permissions,
            };

            if (editMode && editUserRole) {
                // Update existing item
                const webUrl = pageContext.web.absoluteUrl;
                const updateUrl = `${webUrl}/_api/web/lists/getByTitle('SYS_User_Roles')/items(${editUserRole.ID})`;

                const response = await spHttpClient.post(
                    updateUrl,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-HTTP-Method': 'MERGE',
                            'If-Match': '*'
                        },
                        body: JSON.stringify(itemData)
                    }
                );

                if (response.ok) {
                    setState((prev) => ({
                        ...prev,
                        isSubmitting: false,
                        successMessage: 'User role updated successfully!',
                    }));
                } else {
                    throw new Error('Failed to update user role');
                }
            } else {
                // Create new user role
                const userData = {
                    UserId: state.formData.UserId,
                    Role: state.formData.Role,
                    Department: state.formData.Department,
                    Permissions: state.formData.Permissions,
                };
                await sharePointService.createListItem('SYS_User_Roles', userData);
                setState((prev) => ({
                    ...prev,
                    isSubmitting: false,
                    successMessage: 'User role created successfully!',
                    formData: {
                        UserId: undefined,
                        Role: '',
                        Department: '',
                        Permissions: '',
                    },
                }));
            }

            // Close panel after success
            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error
                ? error.message
                : (editMode ? 'Failed to update user role' : 'Failed to create user role');

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, sharePointService, onDismiss, onSubmitSuccess, editMode, editUserRole, spHttpClient, pageContext]);

    const handleCancel = (): void => {
        setState({
            formData: {
                UserId: undefined,
                Role: '',
                Department: '',
                Permissions: '',
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
            headerText={editMode ? 'Edit User Role' : 'New User Role'}
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

                    {/* User Selection - PeoplePicker */}
                    {editMode && editUserRole ? (
                        // Show read-only user info in edit mode
                        <div className={classNames.userInfo}>
                            <Text variant="small" block style={{ color: theme.palette.neutralSecondary }}>
                                User
                            </Text>
                            <Text variant="large" block>
                                {editUserRole.User?.Title || 'Unknown User'}
                            </Text>
                        </div>
                    ) : (
                        // Show PeoplePicker in create mode
                        <div className={classNames.fieldGroup}>
                            <PeoplePicker
                                key="user-role-picker"
                                titleText="Select User"
                                selectedUsers={state.formData.UserId ? [{ id: state.formData.UserId, Title: '', Email: '', LoginName: '' }] : []}
                                personSelectionLimit={1}
                                disabled={isDisabled}
                                spHttpClient={spHttpClient}
                                pageContext={pageContext}
                                webPartContext={webPartContext}
                                onChange={(items) => {
                                    if (items && items.length > 0 && items[0].id) {
                                        handleInputChange('UserId', items[0].id);
                                    } else {
                                        handleInputChange('UserId', undefined);
                                    }
                                }}
                            />
                        </div>
                    )}

                    <Stack tokens={{ childrenGap: 16 }}>
                        {/* Role - Dropdown (Choice column) */}
                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Role"
                                options={roleOptions}
                                selectedKey={state.formData.Role}
                                onChange={(_, option) => handleInputChange('Role', option?.key as string || '')}
                                required
                                disabled={isDisabled}
                                styles={{
                                    dropdown: { width: '100%' },
                                }}
                            />
                        </div>

                        {/* Department - Dropdown (Choice column) */}
                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Department"
                                options={departmentOptions}
                                selectedKey={state.formData.Department}
                                onChange={(_, option) => handleInputChange('Department', option?.key as string || '')}
                                disabled={isDisabled}
                                styles={{
                                    dropdown: { width: '100%' },
                                }}
                            />
                        </div>

                        {/* Permissions - Multi-line text */}
                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Permissions"
                                value={state.formData.Permissions}
                                onChange={(_, value) => handleInputChange('Permissions', value || '')}
                                multiline
                                rows={4}
                                disabled={isDisabled}
                                placeholder="Enter permissions or access levels..."
                            />
                        </div>

                        {/* Action Buttons */}
                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '24px' }}>
                            <DefaultButton
                                text="Cancel"
                                onClick={handleCancel}
                                disabled={isDisabled}
                            />
                            <PrimaryButton
                                text={state.isSubmitting ? 'Saving...' : (editMode ? 'Update' : 'Create')}
                                onClick={handleSubmit}
                                disabled={isDisabled || !state.formData.Role || (!editMode && !state.formData.UserId)}
                            />
                        </div>
                    </Stack>
                </div>
            </div>
        </Panel>
    );
};

export default UserRoleForm;
