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
    ComboBox,
    IComboBoxOption,
    Checkbox,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';
import { IMaterial } from './MaterialsModule';

export interface IMaterialFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmitSuccess?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  editMode?: boolean;
  editMaterial?: IMaterial;
}

interface IMaterialFormData {
  ID?: number;
  Material_Code: string;
  Material_Name: string;
  Category: string;
  UOM: string;
  Standard_Cost: string;
  Active: boolean;
}

interface IFormState {
    formData: IMaterialFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    categoryOptions: IComboBoxOption[];
    uomOptions: IComboBoxOption[];
    showCustomCategory: boolean;
    showCustomUOM: boolean;
}



const MaterialForm: React.FC<IMaterialFormProps> = ({
  isOpen,
  onDismiss,
  onSubmitSuccess,
  spHttpClient,
  pageContext,
  editMode = false,
  editMaterial,
}) => {
    const [state, setState] = React.useState<IFormState>({
        formData: {
            Material_Code: '',
            Material_Name: '',
            Category: '',
            UOM: '',
            Standard_Cost: '',
            Active: true,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        categoryOptions: [],
        uomOptions: [],
        showCustomCategory: false,
        showCustomUOM: false,
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
          const allItems = await sharePointService.getListData('ENT_Materials_Master', undefined, 500, 0, undefined);
          let nextNumber = 1;
          if (allItems && allItems.length > 0) {
            allItems.forEach((item) => {
              if (item.Material_Code) {
                const matches = item.Material_Code.match(/\d+$/);
                if (matches) {
                  const num = parseInt(matches[0], 10);
                  if (num >= nextNumber) {
                    nextNumber = num + 1;
                  }
                }
              }
            });
          }

          const newCode = `MAT-${nextNumber.toString().padStart(4, '0')}`;
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Material_Code: newCode },
          }));
        } catch (error) {
          console.error('Error generating material code:', error);
          // Fallback to timestamp-based code
          const timestamp = Date.now().toString().slice(-6);
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Material_Code: `MAT-${timestamp}` },
          }));
        }
      };

      const loadChoices = async (): Promise<void> => {
        try {
          const [categoryChoices, uomChoices] = await Promise.all([
            sharePointService.getFieldChoices('ENT_Materials_Master', 'Category'),
            sharePointService.getFieldChoices('ENT_Materials_Master', 'UOM'),
          ]);

          const categoryOptions: IComboBoxOption[] = categoryChoices.map(choice => ({
            key: choice,
            text: choice,
          }));
          categoryOptions.push({ key: 'others', text: 'Others (Add new)' });

          const uomOptions: IComboBoxOption[] = uomChoices.map(choice => ({
            key: choice,
            text: choice,
          }));
          uomOptions.push({ key: 'others', text: 'Others (Add new)' });

          setState((prev) => ({
            ...prev,
            categoryOptions,
            uomOptions,
          }));
        } catch (error) {
          console.error('Error loading field choices:', error);
          // Fallback to basic options
          setState((prev) => ({
            ...prev,
            categoryOptions: [{ key: 'others', text: 'Others (Add new)' }],
            uomOptions: [{ key: 'others', text: 'Others (Add new)' }],
          }));
        }
      };

      if (editMode && editMaterial) {
        // Load existing material data for editing
        setState((prev) => ({
          ...prev,
          formData: {
            ID: editMaterial.ID,
            Material_Code: editMaterial.Material_Code,
            Material_Name: editMaterial.Material_Name,
            Category: editMaterial.Category || '',
            UOM: editMaterial.UOM || '',
            Standard_Cost: editMaterial.Standard_Cost?.toString() || '',
            Active: editMaterial.Active,
          },
        }));
      } else {
        // Reset form for new material
        setState((prev) => ({
          ...prev,
          formData: {
            Material_Code: '',
            Material_Name: '',
            Category: '',
            UOM: '',
            Standard_Cost: '',
            Active: true,
          },
        }));

        if (!state.formData.Material_Code) {
          autoGenerateCode().catch(console.error);
        }
      }
      loadChoices().catch(console.error);
    }
  }, [isOpen, editMode, editMaterial, sharePointService]);

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

    const handleInputChange = (fieldName: keyof IMaterialFormData, value: string | boolean | undefined): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [fieldName]: value,
            },
        }));
    };

    const handleComboBoxChange = (fieldName: keyof IMaterialFormData, option?: IComboBoxOption, index?: number, value?: string): void => {
        if (option?.key === 'others') {
            if (fieldName === 'Category') {
                setState((prev) => ({ ...prev, showCustomCategory: true }));
                handleInputChange('Category', '');
            } else if (fieldName === 'UOM') {
                setState((prev) => ({ ...prev, showCustomUOM: true }));
                handleInputChange('UOM', '');
            }
        } else {
            const selectedValue = option ? option.key as string : value || '';
            handleInputChange(fieldName, selectedValue);
            // Hide custom input if a predefined option is selected
            if (fieldName === 'Category') {
                setState((prev) => ({ ...prev, showCustomCategory: false }));
            } else if (fieldName === 'UOM') {
                setState((prev) => ({ ...prev, showCustomUOM: false }));
            }
        }
    };

    const validateForm = (): boolean => {
        if (!state.formData.Material_Code.trim()) {
            setState((prev) => ({ ...prev, error: 'Material Code is required' }));
            return false;
        }

        if (!state.formData.Material_Name.trim()) {
            setState((prev) => ({ ...prev, error: 'Material Name is required' }));
            return false;
        }

        if (!state.formData.Category) {
            setState((prev) => ({ ...prev, error: 'Category is required' }));
            return false;
        }

        if (!state.formData.UOM) {
            setState((prev) => ({ ...prev, error: 'Unit of Measure is required' }));
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
        Title: state.formData.Material_Name,
        Material_Code: state.formData.Material_Code,
        Material_Name: state.formData.Material_Name,
        Category: state.formData.Category,
        UOM: state.formData.UOM,
        Standard_Cost: parseFloat(state.formData.Standard_Cost) || 0,
        Active: state.formData.Active,
      };

      if (editMode && state.formData.ID) {
        // Update existing material
        await sharePointService.updateListItem('ENT_Materials_Master', state.formData.ID, itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Material updated successfully!',
        }));
      } else {
        // Create new material
        await sharePointService.createListItem('ENT_Materials_Master', itemData);

        // Try to add new choices if they don't exist
        const addChoicesPromises = [];
        if (state.formData.Category && !state.categoryOptions.some(opt => opt.key === state.formData.Category)) {
          addChoicesPromises.push(sharePointService.addFieldChoice('ENT_Materials_Master', 'Category', state.formData.Category));
        }
        if (state.formData.UOM && !state.uomOptions.some(opt => opt.key === state.formData.UOM)) {
          addChoicesPromises.push(sharePointService.addFieldChoice('ENT_Materials_Master', 'UOM', state.formData.UOM));
        }
        await Promise.all(addChoicesPromises);

        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Material created successfully!',
          formData: {
            Material_Code: '',
            Material_Name: '',
            Category: '',
            UOM: '',
            Standard_Cost: '',
            Active: true,
          },
          showCustomCategory: false,
          showCustomUOM: false,
        }));
      }

      setTimeout(() => {
        onDismiss();
        if (onSubmitSuccess) {
          onSubmitSuccess();
        }
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : `Failed to ${editMode ? 'update' : 'create'} material`;
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [state.formData, editMode, sharePointService, onDismiss, onSubmitSuccess]);

    const handleCancel = (): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                Material_Code: '',
                Material_Name: '',
                Category: '',
                UOM: '',
                Standard_Cost: '',
                Active: true,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            showCustomCategory: false,
            showCustomUOM: false,
        }));
        onDismiss();
    };

    const isDisabled = state.isSubmitting;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText={editMode ? "Edit Material" : "Add Material"}
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
                                label="Material Code"
                                placeholder="Auto-generated"
                                value={state.formData.Material_Code}
                                disabled={true}
                                required
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Material Name"
                                placeholder="Enter material name"
                                value={state.formData.Material_Name}
                                onChange={(_, value) => handleInputChange('Material_Name', value)}
                                disabled={isDisabled}
                                required
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <ComboBox
                                label="Category"
                                placeholder="Select category"
                                allowFreeform={false}
                                autoComplete="on"
                                options={state.categoryOptions}
                                selectedKey={state.showCustomCategory ? undefined : state.formData.Category}
                                onChange={(event, option, index, value) => handleComboBoxChange('Category', option, index, value)}
                                disabled={isDisabled}
                                required
                            />
                            {state.showCustomCategory && (
                                <TextField
                                    placeholder="Enter new category"
                                    value={state.formData.Category}
                                    onChange={(_, value) => handleInputChange('Category', value)}
                                    disabled={isDisabled}
                                    required
                                />
                            )}
                        </div>

                        <div className={classNames.fieldGroup}>
                            <ComboBox
                                label="Unit of Measure"
                                placeholder="Select UOM"
                                allowFreeform={false}
                                autoComplete="on"
                                options={state.uomOptions}
                                selectedKey={state.showCustomUOM ? undefined : state.formData.UOM}
                                onChange={(event, option, index, value) => handleComboBoxChange('UOM', option, index, value)}
                                disabled={isDisabled}
                                required
                            />
                            {state.showCustomUOM && (
                                <TextField
                                    placeholder="Enter new unit of measure"
                                    value={state.formData.UOM}
                                    onChange={(_, value) => handleInputChange('UOM', value)}
                                    disabled={isDisabled}
                                    required
                                />
                            )}
                        </div>

                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Standard Cost"
                                placeholder="Enter standard cost"
                                type="number"
                                value={state.formData.Standard_Cost}
                                onChange={(_, value) => handleInputChange('Standard_Cost', value)}
                                disabled={isDisabled}
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <Checkbox
                                label="Active"
                                checked={state.formData.Active}
                                onChange={(_, checked) => handleInputChange('Active', checked)}
                                disabled={isDisabled}
                            />
                        </div>

                        <div className={classNames.buttonGroup}>
                            <PrimaryButton
                              text={editMode ? "Update Material" : "Add Material"}
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
                                <Spinner size={SpinnerSize.medium} label="Creating material..." />
                            </div>
                        )}
                    </Stack>
                </div>
            </div>
        </Panel>
    );
};

export default MaterialForm;
