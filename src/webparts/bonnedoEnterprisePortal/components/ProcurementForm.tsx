import * as React from 'react';
import {
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  mergeStyleSets,
  Panel,
  PanelType,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IProcurementFormData } from '../models/DataModels';
//import styles from './ProcurementForm.module.scss';

export interface IProcurementFormProps {
  formType: 'MR' | 'PR' | 'PO' | 'GRN';
  listName: string;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  onSubmitSuccess?: (itemId: number) => void;
  onCancel?: () => void;
  isOpen?: boolean;
  onDismiss?: () => void;
}

interface IFormState {
  formData: IProcurementFormData;
  isSubmitting: boolean;
  error: string | undefined;
  successMessage: string | undefined;
}

const UOM_OPTIONS: IDropdownOption[] = [
  { key: 'PCS', text: 'Pieces' },
  { key: 'KG', text: 'Kilograms' },
  { key: 'L', text: 'Liters' },
  { key: 'M', text: 'Meters' },
  { key: 'BOX', text: 'Box' },
  { key: 'PACK', text: 'Pack' },
  { key: 'ROLL', text: 'Roll' },
  { key: 'SHEET', text: 'Sheet' },
];

const ProcurementForm: React.FC<IProcurementFormProps> = ({
  formType,
  listName,
  spHttpClient,
  pageContext,
  onSubmitSuccess,
  onCancel,
  isOpen,
  onDismiss,
}) => {
  const [state, setState] = React.useState<IFormState>({
    formData: {
      Project_Code: '',
      Quantity: 0,
      UOM: 'PCS',
    },
    isSubmitting: false,
    error: undefined,
    successMessage: undefined,
  });

  const classNames = mergeStyleSets({
    formContainer: {
      padding: '20px',
      maxWidth: '600px',
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
  const handleInputChange = (fieldName: string, value: any): void => {
    setState((prev) => ({
      ...prev,
      formData: {
        ...prev.formData,
        [fieldName]: value,
      },
    }));
  };

  const validateForm = (): boolean => {
    if (!state.formData.Project_Code.trim()) {
      setState((prev) => ({ ...prev, error: 'Project Code is required' }));
      return false;
    }

    if (state.formData.Quantity <= 0) {
      setState((prev) => ({ ...prev, error: 'Quantity must be greater than 0' }));
      return false;
    }

    if (formType === 'MR' && !state.formData.Material?.trim()) {
      setState((prev) => ({ ...prev, error: 'Material is required' }));
      return false;
    }

    if (formType === 'PR' && !state.formData.Description?.trim()) {
      setState((prev) => ({ ...prev, error: 'Description is required' }));
      return false;
    }

    if (formType === 'PO' && !state.formData.Vendor?.trim()) {
      setState((prev) => ({ ...prev, error: 'Vendor is required' }));
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const itemData: any = {
        Title: `${formType}-${Date.now()}`,
        Project_Code: state.formData.Project_Code,
        Quantity: state.formData.Quantity,
        UOM: state.formData.UOM,
        Status: 'Draft',
        Approval_Status: 'Pending',
      };

      // Add form-specific fields
      if (formType === 'MR') {
        itemData.Material = state.formData.Material;
        itemData.Request_Date = new Date().toISOString();
      } else if (formType === 'PR') {
        itemData.Description = state.formData.Description;
        itemData.EstimatedCost = state.formData.EstimatedCost || 0;
        itemData.Request_Date = new Date().toISOString();
      } else if (formType === 'PO') {
        itemData.Vendor = state.formData.Vendor;
        itemData.UnitPrice = state.formData.UnitPrice || 0;
        itemData.TotalAmount = (state.formData.UnitPrice || 0) * state.formData.Quantity;
        itemData.DeliveryDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
      } else if (formType === 'GRN') {
        itemData.PO_Number = state.formData.Material;
        itemData.Vendor = state.formData.Vendor;
        itemData.Quantity_Received = state.formData.Quantity;
        itemData.Received_Date = new Date().toISOString();
      }

      // Create item in SharePoint
      const webUrl = pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getByTitle('${listName}')/items`;

      const response = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(itemData),
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to create item: ${response.statusText}`);
      }

      const createdItem = await response.json();

      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        successMessage: `${formType} created successfully!`,
        formData: {
          Project_Code: '',
          Quantity: 0,
          UOM: 'PCS',
        },
      }));

      // Notify parent component
      if (onSubmitSuccess) {
        setTimeout(() => {
          onSubmitSuccess(createdItem.ID);
        }, 1500);
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to create item';
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [state.formData, formType, listName, spHttpClient, pageContext, onSubmitSuccess]);

  const isDisabled = state.isSubmitting;

  const formContent = (
    <div className={classNames.formContainer} style={{ paddingTop: '20px' }}>
      {state.error && (
        <div style={{ marginBottom: '16px', marginTop: '8px' }}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {state.error}
          </MessageBar>
        </div>
      )}

      {state.successMessage && (
        <div style={{ marginBottom: '16px', marginTop: '8px' }}>
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
            placeholder="Enter project code"
            value={state.formData.Project_Code}
            onChange={(_, value) => handleInputChange('Project_Code', value)}
            disabled={isDisabled}
            required
          />
        </div>

        {/* Material (for MR) */}
        {formType === 'MR' && (
          <div className={classNames.fieldGroup}>
            <TextField
              label="Material"
              placeholder="Enter material name"
              value={state.formData.Material || ''}
              onChange={(_, value) => handleInputChange('Material', value)}
              disabled={isDisabled}
              required
            />
          </div>
        )}

        {/* Description (for PR) */}
        {formType === 'PR' && (
          <div className={classNames.fieldGroup}>
            <TextField
              label="Description"
              placeholder="Enter description"
              value={state.formData.Description || ''}
              onChange={(_, value) => handleInputChange('Description', value)}
              disabled={isDisabled}
              multiline
              rows={3}
              required
            />
          </div>
        )}

        {/* Vendor (for PO and GRN) */}
        {(formType === 'PO' || formType === 'GRN') && (
          <div className={classNames.fieldGroup}>
            <TextField
              label="Vendor"
              placeholder="Enter vendor name"
              value={state.formData.Vendor || ''}
              onChange={(_, value) => handleInputChange('Vendor', value)}
              disabled={isDisabled}
              required
            />
          </div>
        )}

        {/* Quantity */}
        <div className={classNames.fieldGroup}>
          <TextField
            label="Quantity"
            placeholder="Enter quantity"
            type="number"
            value={state.formData.Quantity.toString()}
            onChange={(_, value) => handleInputChange('Quantity', parseInt(value || '0', 10))}
            disabled={isDisabled}
            required
          />
        </div>

        {/* UOM */}
        <div className={classNames.fieldGroup}>
          <Dropdown
            label="Unit of Measure"
            options={UOM_OPTIONS}
            selectedKey={state.formData.UOM}
            onChange={(_, option) => handleInputChange('UOM', option?.key)}
            disabled={isDisabled}
          />
        </div>

        {/* Unit Price (for PO) */}
        {formType === 'PO' && (
          <div className={classNames.fieldGroup}>
            <TextField
              label="Unit Price"
              placeholder="Enter unit price"
              type="number"
              value={state.formData.UnitPrice?.toString() || ''}
              onChange={(_, value) => handleInputChange('UnitPrice', parseFloat(value || '0'))}
              disabled={isDisabled}
            />
          </div>
        )}

        {/* Estimated Cost (for PR) */}
        {formType === 'PR' && (
          <div className={classNames.fieldGroup}>
            <TextField
              label="Estimated Cost"
              placeholder="Enter estimated cost"
              type="number"
              value={state.formData.EstimatedCost?.toString() || ''}
              onChange={(_, value) => handleInputChange('EstimatedCost', parseFloat(value || '0'))}
              disabled={isDisabled}
            />
          </div>
        )}

        {/* Notes */}
        <div className={classNames.fieldGroup}>
          <TextField
            label="Notes"
            placeholder="Enter any additional notes"
            value={state.formData.Notes || ''}
            onChange={(_, value) => handleInputChange('Notes', value)}
            disabled={isDisabled}
            multiline
            rows={2}
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
            onClick={onCancel}
            disabled={isDisabled}
            iconProps={{ iconName: 'Cancel' }}
          />
        </div>

        {state.isSubmitting && (
          <div style={{ textAlign: 'center', marginTop: '16px' }}>
            <Spinner size={SpinnerSize.medium} label="Submitting..." />
          </div>
        )}
      </Stack>
    </div>
  );

  // If isOpen is provided, render as a side panel
  if (isOpen !== undefined && onDismiss) {
    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText={`New ${formType === 'MR' ? 'Material Request' : formType === 'PR' ? 'Purchase Requisition' : formType === 'PO' ? 'Purchase Order' : 'GRN'}`}
        closeButtonAriaLabel="Close"
      >
        <div style={{ padding: '0 20px', height: '100%', overflowY: 'auto' }}>
          {formContent}
        </div>
      </Panel>
    );
  }

  return formContent;
};

export default ProcurementForm;
