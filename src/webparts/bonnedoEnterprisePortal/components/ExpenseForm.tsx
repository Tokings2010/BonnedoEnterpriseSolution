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
    ComboBox,
    IComboBoxOption,
    getTheme,
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService, IListItem } from '../services/SharePointService';

export interface IExpenseFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

interface IFormData {
    Project_CodeId: number | undefined;
    Expense_Type: string;
    Amount: number | undefined;
    Recorded_ById: number | undefined;
    Notes: string;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
    expenseTypeOptions: IComboBoxOption[];
    showCustomExpenseType: boolean;
}



const ExpenseForm: React.FC<IExpenseFormProps> = ({
    isOpen,
    onDismiss,
    onSubmitSuccess,
    spHttpClient,
    pageContext,
    webPartContext,
}) => {
    const theme = getTheme();
    const [state, setState] = React.useState<IFormState>({
        formData: {
            Project_CodeId: undefined,
            Expense_Type: 'Other',
            Amount: undefined,
            Recorded_ById: undefined,
            Notes: '',
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        isLoadingLookups: true,
        currentUserId: undefined,
        expenseTypeOptions: [],
        showCustomExpenseType: false,
    });

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    // Load lookup data on mount
    React.useEffect(() => {
        const loadLookups = async (): Promise<void> => {
            try {
                const [projects, currentUser, expenseTypeChoices] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getCurrentUser(),
                    sharePointService.getFieldChoices('FIN_Expense_Register', 'Expense_Type'),
                ]);

                const expenseTypeOptions: IComboBoxOption[] = expenseTypeChoices.map(choice => ({
                    key: choice,
                    text: choice,
                }));
                expenseTypeOptions.push({ key: 'others', text: 'Others (Add new)' });

                setState((prev) => ({
                    ...prev,
                    projects,
                    currentUserId: currentUser?.ID,
                    expenseTypeOptions,
                    formData: {
                        ...prev.formData,
                        // Auto-populate Recorded By with current user
                        Recorded_ById: currentUser?.ID,
                    },
                    isLoadingLookups: false,
                }));
            } catch (error) {
                console.error('Error loading lookups:', error);
                // Fallback to basic options
                setState((prev) => ({
                    ...prev,
                    expenseTypeOptions: [{ key: 'others', text: 'Others (Add new)' }],
                    isLoadingLookups: false,
                    error: 'Failed to load lookup data',
                }));
            }
        };

        if (isOpen) {
            loadLookups();
        }
    }, [isOpen, sharePointService]);

    const classNames = mergeStyleSets({
        formContainer: {
            padding: '20px',
        },
        fieldGroup: {
            marginBottom: '16px',
        },
        sectionHeader: {
            fontSize: '16px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
            marginBottom: '16px',
            marginTop: '24px',
            borderBottom: `1px solid ${theme.palette.neutralLight}`,
            paddingBottom: '8px',
        },
        buttonGroup: {
            display: 'flex',
            gap: '12px',
            marginTop: '24px',
            flexWrap: 'wrap',
        },
    });

    const handleInputChange = (fieldName: keyof IFormData, value: string | number | undefined): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [fieldName]: value,
            },
        }));
    };

    const handleExpenseTypeChange = (option?: IComboBoxOption, index?: number, value?: string): void => {
        if (option?.key === 'others') {
            setState((prev) => ({ ...prev, showCustomExpenseType: true }));
            handleInputChange('Expense_Type', '');
        } else {
            const selectedValue = option ? option.key as string : value || '';
            handleInputChange('Expense_Type', selectedValue);
            // Hide custom input if a predefined option is selected
            setState((prev) => ({ ...prev, showCustomExpenseType: false }));
        }
    };

    const validateForm = (): boolean => {
        if (!state.formData.Project_CodeId) {
            setState((prev) => ({ ...prev, error: 'Project Code is required' }));
            return false;
        }
        if (!state.formData.Expense_Type) {
            setState((prev) => ({ ...prev, error: 'Expense Type is required' }));
            return false;
        }
        if (!state.formData.Amount) {
            setState((prev) => ({ ...prev, error: 'Amount is required' }));
            return false;
        }
        if (!state.formData.Recorded_ById) {
            setState((prev) => ({ ...prev, error: 'Recorded By is required' }));
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
                Project_CodeId: state.formData.Project_CodeId,
                Expense_Type: state.formData.Expense_Type,
                Amount: state.formData.Amount || 0,
                Recorded_ById: state.formData.Recorded_ById,
                Notes: state.formData.Notes,
                Approval_Status: 'Pending',
            };

            await sharePointService.createListItem('FIN_Expense_Register', itemData);

            // Try to add new expense type if it's not in the existing choices
            if (state.formData.Expense_Type && !state.expenseTypeOptions.some(opt => opt.key === state.formData.Expense_Type)) {
                await sharePointService.addFieldChoice('FIN_Expense_Register', 'Expense_Type', state.formData.Expense_Type);
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Expense created successfully!',
                formData: {
                    Project_CodeId: undefined,
                    Expense_Type: 'Other',
                    Amount: undefined,
                    Recorded_ById: prev.currentUserId,
                    Notes: '',
                },
                showCustomExpenseType: false,
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Expense';
            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, sharePointService, onDismiss, onSubmitSuccess]);

    const handleCancel = (): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                Project_CodeId: undefined,
                Expense_Type: 'Other',
                Amount: undefined,
                Recorded_ById: prev.currentUserId,
                Notes: '',
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            showCustomExpenseType: false,
        }));
        onDismiss();
    };

    // Convert lookup items to dropdown options
    const projectOptions: IDropdownOption[] = state.projects.map((item) => ({
        key: item.Id,
        text: item.Project_Code || `Project ${item.Id}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Expense"
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

                    {state.isLoadingLookups && (
                        <div style={{ textAlign: 'center', padding: '20px' }}>
                            <Spinner size={SpinnerSize.medium} label="Loading data..." />
                        </div>
                    )}

                    {!state.isLoadingLookups && (
                        <Stack tokens={{ childrenGap: 16 }}>
                            {/* Section 1: Expense Details */}
                            <div className={classNames.sectionHeader}>Expense Details</div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Project Code"
                                    options={projectOptions}
                                    selectedKey={state.formData.Project_CodeId}
                                    onChange={(e, option) => {
                                        handleInputChange('Project_CodeId', option?.key as number);
                                    }}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Project"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <ComboBox
                                    label="Expense Type"
                                    placeholder="Select expense type"
                                    allowFreeform={false}
                                    autoComplete="on"
                                    options={state.expenseTypeOptions}
                                    selectedKey={state.showCustomExpenseType ? undefined : state.formData.Expense_Type}
                                    onChange={(event, option, index, value) => handleExpenseTypeChange(option, index, value)}
                                    disabled={isDisabled}
                                    required
                                />
                                {state.showCustomExpenseType && (
                                    <TextField
                                        placeholder="Enter new expense type"
                                        value={state.formData.Expense_Type}
                                        onChange={(_, value) => handleInputChange('Expense_Type', value)}
                                        disabled={isDisabled}
                                        required
                                    />
                                )}
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Amount"
                                    type="number"
                                    value={state.formData.Amount?.toString() || ''}
                                    onChange={(e, value) => handleInputChange('Amount', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    required
                                    prefix="₦"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Notes"
                                    value={state.formData.Notes}
                                    onChange={(e, value) => handleInputChange('Notes', value || '')}
                                    disabled={isDisabled}
                                    multiline
                                    rows={3}
                                />
                            </div>

                            {/* Section 2: Personnel */}
                            <div className={classNames.sectionHeader}>Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="recorded-by-picker"
                                    titleText="Recorded By"
                                    selectedUsers={state.formData.Recorded_ById ? [{ id: state.formData.Recorded_ById, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    required={true}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Recorded_ById', items[0].id);
                                        } else {
                                            handleInputChange('Recorded_ById', undefined);
                                        }
                                    }}
                                />
                            </div>

                            {/* Buttons */}
                            <div className={classNames.buttonGroup}>
                                <PrimaryButton
                                    text="Submit"
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
                        </Stack>
                    )}
                </div>
            </div>
        </Panel>
    );
};

export default ExpenseForm;
