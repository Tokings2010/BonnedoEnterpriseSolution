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
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';
import type { IWarehouse } from './WarehousesModule';

export interface IWarehouseFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmitSuccess?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  editMode?: boolean;
  editWarehouse?: IWarehouse;
}

interface IWarehouseFormData {
  ID?: number;
  WarehouseCode: string;
  Warehouse_Name: string;
  Location: string;
  Status: string;
}

interface IFormState {
  formData: IWarehouseFormData;
  isSubmitting: boolean;
  error: string | undefined;
  successMessage: string | undefined;
}

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Active', text: 'Active' },
  { key: 'Inactive', text: 'Inactive' },
];

const WarehouseForm: React.FC<IWarehouseFormProps> = ({
  isOpen,
  onDismiss,
  onSubmitSuccess,
  spHttpClient,
  pageContext,
  editMode = false,
  editWarehouse,
}) => {
  const [state, setState] = React.useState<IFormState>({
    formData: {
      WarehouseCode: '',
      Warehouse_Name: '',
      Location: '',
      Status: 'Active',
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
    if (!isOpen) {
      return;
    }

    const autoGenerateCode = async (): Promise<void> => {
      try {
        const allItems = await sharePointService.getListData(
          SHAREPOINT_LISTS.WAREHOUSES_MASTER,
          undefined,
          500,
          0,
          undefined
        );
        let nextNumber = 1;
        allItems.forEach((item) => {
          if (item.WarehouseCode) {
            const matches = String(item.WarehouseCode).match(/\d+$/);
            if (matches) {
              const num = parseInt(matches[0], 10);
              if (num >= nextNumber) {
                nextNumber = num + 1;
              }
            }
          }
        });
        const newCode = `WH-${nextNumber.toString().padStart(4, '0')}`;
        setState((prev) => ({
          ...prev,
          formData: { ...prev.formData, WarehouseCode: newCode },
        }));
      } catch (error) {
        console.error('Error generating warehouse code:', error);
        const timestamp = Date.now().toString().slice(-6);
        setState((prev) => ({
          ...prev,
          formData: { ...prev.formData, WarehouseCode: `WH-${timestamp}` },
        }));
      }
    };

    if (editMode && editWarehouse) {
      setState((prev) => ({
        ...prev,
        formData: {
          ID: editWarehouse.ID,
          WarehouseCode: editWarehouse.WarehouseCode,
          Warehouse_Name: editWarehouse.Warehouse_Name,
          Location: editWarehouse.Location || '',
          Status: editWarehouse.Status,
        },
        error: undefined,
        successMessage: undefined,
      }));
    } else {
      setState((prev) => ({
        ...prev,
        formData: {
          WarehouseCode: '',
          Warehouse_Name: '',
          Location: '',
          Status: 'Active',
        },
        error: undefined,
        successMessage: undefined,
      }));
      autoGenerateCode().catch(console.error);
    }
  }, [isOpen, editMode, editWarehouse, sharePointService]);

  const classNames = mergeStyleSets({
    formContainer: { padding: '20px' },
    fieldGroup: { marginBottom: '16px' },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
      flexWrap: 'wrap',
    },
  });

  const handleInputChange = (fieldName: keyof IWarehouseFormData, value: string | undefined): void => {
    setState((prev) => ({
      ...prev,
      formData: { ...prev.formData, [fieldName]: value },
    }));
  };

  const validateForm = (): boolean => {
    if (!state.formData.WarehouseCode.trim()) {
      setState((prev) => ({ ...prev, error: 'Warehouse Code is required' }));
      return false;
    }
    if (!state.formData.Warehouse_Name.trim()) {
      setState((prev) => ({ ...prev, error: 'Warehouse Name is required' }));
      return false;
    }
    if (!state.formData.Status) {
      setState((prev) => ({ ...prev, error: 'Status is required' }));
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
        Title: state.formData.WarehouseCode,
        WarehouseCode: state.formData.WarehouseCode,
        Warehouse_Name: state.formData.Warehouse_Name,
        Location: state.formData.Location,
        Status: state.formData.Status,
      };

      if (editMode && state.formData.ID) {
        await sharePointService.updateListItem(
          SHAREPOINT_LISTS.WAREHOUSES_MASTER,
          state.formData.ID,
          itemData
        );
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Warehouse updated successfully!',
        }));
      } else {
        await sharePointService.createListItem(SHAREPOINT_LISTS.WAREHOUSES_MASTER, itemData);
        setState((prev) => ({
          ...prev,
          isSubmitting: false,
          successMessage: 'Warehouse created successfully!',
        }));
      }

      setTimeout(() => {
        onDismiss();
        onSubmitSuccess?.();
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error
        ? error.message
        : `Failed to ${editMode ? 'update' : 'create'} warehouse`;
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
        WarehouseCode: '',
        Warehouse_Name: '',
        Location: '',
        Status: 'Active',
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
      headerText={editMode ? 'Edit Warehouse' : 'Add Warehouse'}
      closeButtonAriaLabel="Close"
      isLightDismiss={!state.isSubmitting}
    >
      <div style={{ padding: '0 20px', height: '100%', overflowY: 'auto' }}>
        <div className={classNames.formContainer}>
          {state.error && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline>
              {state.error}
            </MessageBar>
          )}
          {state.successMessage && (
            <MessageBar messageBarType={MessageBarType.success} isMultiline>
              {state.successMessage}
            </MessageBar>
          )}

          <Stack tokens={{ childrenGap: 16 }}>
            <TextField
              label="Warehouse Code"
              value={state.formData.WarehouseCode}
              disabled
              required
            />
            <TextField
              label="Warehouse Name"
              placeholder="Enter warehouse name"
              value={state.formData.Warehouse_Name}
              onChange={(_, value) => handleInputChange('Warehouse_Name', value)}
              disabled={isDisabled}
              required
            />
            <TextField
              label="Physical Location"
              placeholder="Enter address or site description"
              value={state.formData.Location}
              onChange={(_, value) => handleInputChange('Location', value)}
              disabled={isDisabled}
            />
            <Dropdown
              label="Status"
              options={STATUS_OPTIONS}
              selectedKey={state.formData.Status}
              onChange={(_, option) => handleInputChange('Status', option?.key as string)}
              disabled={isDisabled}
              required
            />
            <div className={classNames.buttonGroup}>
              <PrimaryButton
                text={editMode ? 'Update Warehouse' : 'Add Warehouse'}
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
              <Spinner size={SpinnerSize.medium} label={editMode ? 'Updating warehouse...' : 'Creating warehouse...'} />
            )}
          </Stack>
        </div>
      </div>
    </Panel>
  );
};

export default WarehouseForm;
