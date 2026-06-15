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
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';
import { IVendor } from './VendorsModule';

export interface IVendorFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmitSuccess?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  editMode?: boolean;
  editVendor?: IVendor;
}

interface IVendorFormData {
  ID?: number;
  Vendor_Code: string;
  Vendor_Name: string;
  Vendor_Category: string;
  Vendor_Status: string;
}

interface IFormState {
    formData: IVendorFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
}

const CATEGORY_OPTIONS: IDropdownOption[] = [
    { key: 'Services', text: 'Services' },
    { key: 'Goods', text: 'Goods' },
    { key: 'Works', text: 'Works' },
    { key: 'IT Services', text: 'IT Services' },
    { key: 'Consulting', text: 'Consulting' },
    { key: 'Logistics', text: 'Logistics' },
    { key: 'Other', text: 'Other' },
];

const STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'Active', text: 'Active' },
    { key: 'Inactive', text: 'Inactive' },
    { key: 'Pending', text: 'Pending' },
    { key: 'Blocked', text: 'Blocked' },
];

const VendorForm: React.FC<IVendorFormProps> = ({
  isOpen,
  onDismiss,
  onSubmitSuccess,
  spHttpClient,
  pageContext,
  editMode = false,
  editVendor,
}) => {
    const [state, setState] = React.useState<IFormState>({
        formData: {
            Vendor_Code: '',
            Vendor_Name: '',
            Vendor_Category: '',
            Vendor_Status: 'Active',
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
      const autoGenerateCode = async (): Promise<void> => {
        try {
          // Get all items and find the highest number
          const allItems = await sharePointService.getListData('ENT_Vendors_Master', undefined, 500, 0, undefined);
          let nextNumber = 1;
          if (allItems && allItems.length > 0) {
            allItems.forEach((item) => {
              if (item.Vendor_Code) {
                const matches = item.Vendor_Code.match(/\d+$/);
                if (matches) {
                  const num = parseInt(matches[0], 10);
                  if (num >= nextNumber) {
                    nextNumber = num + 1;
                  }
                }
              }
            });
          }

          const newCode = `VND-${nextNumber.toString().padStart(4, '0')}`;
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Vendor_Code: newCode },
          }));
        } catch (error) {
          console.error('Error generating vendor code:', error);
          // Fallback to timestamp-based code
          const timestamp = Date.now().toString().slice(-6);
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Vendor_Code: `VND-${timestamp}` },
          }));
        }
      };

      if (editMode && editVendor) {
        // Load existing vendor data for editing
        setState((prev) => ({
          ...prev,
          formData: {
            ID: editVendor.ID,
            Vendor_Code: editVendor.Vendor_Code,
            Vendor_Name: editVendor.Vendor_Name,
            Vendor_Category: editVendor.Vendor_Category,
            Vendor_Status: editVendor.Vendor_Status,
          },
        }));
      } else {
        // Reset form for new vendor
        setState((prev) => ({
          ...prev,
          formData: {
            Vendor_Code: '',
            Vendor_Name: '',
            Vendor_Category: '',
            Vendor_Status: 'Active',
          },
        }));

        if (!state.formData.Vendor_Code) {
          autoGenerateCode().catch(console.error);
        }
      }
    }
  }, [isOpen, editMode, editVendor, sharePointService]);

    const classNames = mergeStyleSets({
        formContainer: {
            padding: '20px',
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

    const handleInputChange = (fieldName: keyof IVendorFormData, value: string | undefined): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [fieldName]: value,
            },
        }));
    };

    const validateForm = (): boolean => {
        if (!state.formData.Vendor_Code.trim()) {
            setState((prev) => ({ ...prev, error: 'Vendor Code is required' }));
            return false;
        }

        if (!state.formData.Vendor_Name.trim()) {
            setState((prev) => ({ ...prev, error: 'Vendor Name is required' }));
            return false;
        }

        if (!state.formData.Vendor_Category) {
            setState((prev) => ({ ...prev, error: 'Vendor Category is required' }));
            return false;
        }

        if (!state.formData.Vendor_Status) {
            setState((prev) => ({ ...prev, error: 'Vendor Status is required' }));
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
        Title: state.formData.Vendor_Name,
        Vendor_Code: state.formData.Vendor_Code,
        Vendor_Name: state.formData.Vendor_Name,
        Vendor_Category: state.formData.Vendor_Category,
        Vendor_Status: state.formData.Vendor_Status,
      };

      if (editMode && state.formData.ID) {
        // Update existing vendor
        await sharePointService.updateListItem('ENT_Vendors_Master', state.formData.ID, itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Vendor updated successfully!',
        }));
      } else {
        // Create new vendor
        await sharePointService.createListItem('ENT_Vendors_Master', itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Vendor created successfully!',
          formData: {
            Vendor_Code: '',
            Vendor_Name: '',
            Vendor_Category: '',
            Vendor_Status: 'Active',
          },
        }));
      }

      setTimeout(() => {
        onDismiss();
        if (onSubmitSuccess) {
          onSubmitSuccess();
        }
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : `Failed to ${editMode ? 'update' : 'create'} vendor`;
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [state.formData, editMode, sharePointService, onDismiss, onSubmitSuccess]);

    const handleCancel = (): void => {
        setState({
            formData: {
                Vendor_Code: '',
                Vendor_Name: '',
                Vendor_Category: '',
                Vendor_Status: 'Active',
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
            headerText={editMode ? "Edit Vendor" : "Add Vendor"}
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
                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Vendor Code"
                                placeholder="Auto-generated"
                                value={state.formData.Vendor_Code}
                                disabled={true}
                                required
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Vendor Name"
                                placeholder="Enter vendor name"
                                value={state.formData.Vendor_Name}
                                onChange={(_, value) => handleInputChange('Vendor_Name', value)}
                                disabled={isDisabled}
                                required
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Vendor Category"
                                placeholder="Select category"
                                options={CATEGORY_OPTIONS}
                                selectedKey={state.formData.Vendor_Category}
                                onChange={(_, option) => handleInputChange('Vendor_Category', option?.key as string)}
                                disabled={isDisabled}
                                required
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Vendor Status"
                                placeholder="Select status"
                                options={STATUS_OPTIONS}
                                selectedKey={state.formData.Vendor_Status}
                                onChange={(_, option) => handleInputChange('Vendor_Status', option?.key as string)}
                                disabled={isDisabled}
                                required
                            />
                        </div>

                        <div className={classNames.buttonGroup}>
                            <PrimaryButton
                              text={editMode ? "Update Vendor" : "Add Vendor"}
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
                                <Spinner size={SpinnerSize.medium} label="Creating vendor..." />
                            </div>
                        )}
                    </Stack>
                </div>
            </div>
        </Panel>
    );
};

export default VendorForm;
