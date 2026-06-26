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
import { QRCodeService } from '../services/QRCodeService';
import { IMaterial } from './MaterialsModule';

// ─── Category Code Mapping (matches HTML reference design) ──────────────────
const CATEGORY_CODES: Record<string, string> = {
  'Pipes': 'PIP',
  'Valves': 'VAL',
  'Fittings': 'FIT',
  'Valves, Fittings & Accessories': 'FIT',
  'Reducers': 'RED',
  'Flanges': 'FLG',
  'Tees': 'TEE',
  'Bolts': 'BLT',
  'Nuts': 'NUT',
  'Gaskets': 'GSK',
  'OEM Hardware & Software': 'OEM',
  'OEM': 'OEM',
  'Consumables': 'CON',
  'Consumable': 'CON',
  'PPE': 'PPE',
  'Electrical': 'ELE',
  'Instruments': 'INS',
  'Safety': 'SPE',
};

// ─── Sub-Type Code Mapping ──────────────────────────────────────────────────
const SUBTYPE_CODES: Record<string, string> = {
  'Carbon Steel Line': 'CSL',
  'Stainless Steel': 'SST',
  'GRE': 'GRE',
  'API 600': 'GAT',
  'Class 150': 'CLS',
  'Long Radius BW': 'ELB',
  'Concentric BW': 'CON',
  '150# RF': 'WNK',
  'Equal BW': 'EQL',
  'Grade 8.8': 'HEX',
  'OEM Module': 'PLG',
  'PPE': 'HLM',
  'Consumable': 'WRD',
  'N/A': '00',
};

function getCategoryCode(category: string): string {
  return CATEGORY_CODES[category] || category.substring(0, 3).toUpperCase();
}

function getSubTypeCode(subType: string): string {
  return SUBTYPE_CODES[subType] || subType.substring(0, 3).toUpperCase();
}

function getSizeCode(size: string): string {
  if (!size || size === 'N/A') return '00';
  // Extract numeric portion: "6"" -> "06", "10"" -> "10", "M16" -> "M16", "3.2mm" -> "3M"
  const cleaned = size.replace(/"/g, '').trim();
  // If it starts with a letter (like M16), keep it
  if (/^[A-Z]/i.test(cleaned)) return cleaned;
  // If it's a number with decimal, take integer part
  const num = parseFloat(cleaned);
  if (isNaN(num)) return '00';
  return Math.floor(num).toString().padStart(2, '0');
}

const DEFAULT_SUB_TYPE_OPTIONS: IComboBoxOption[] = [
  { key: 'Carbon Steel Line', text: 'Carbon Steel Line' },
  { key: 'Stainless Steel', text: 'Stainless Steel' },
  { key: 'GRE', text: 'GRE' },
  { key: 'API 600', text: 'API 600' },
  { key: 'Class 150', text: 'Class 150' },
  { key: 'Long Radius BW', text: 'Long Radius BW' },
  { key: 'Concentric BW', text: 'Concentric BW' },
  { key: '150# RF', text: '150# RF' },
  { key: 'Equal BW', text: 'Equal BW' },
  { key: 'Grade 8.8', text: 'Grade 8.8' },
  { key: 'OEM Module', text: 'OEM Module' },
  { key: 'Consumable', text: 'Consumable' },
  { key: 'PPE', text: 'PPE' },
  { key: 'N/A', text: 'N/A' },
];

const DEFAULT_SIZE_OPTIONS: IComboBoxOption[] = [
  { key: '2"', text: '2"' },
  { key: '3"', text: '3"' },
  { key: '4"', text: '4"' },
  { key: '6"', text: '6"' },
  { key: '8"', text: '8"' },
  { key: '10"', text: '10"' },
  { key: '12"', text: '12"' },
  { key: '14"', text: '14"' },
  { key: '16"', text: '16"' },
  { key: 'M16', text: 'M16' },
  { key: '3.2mm', text: '3.2mm' },
  { key: 'N/A', text: 'N/A' },
];

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
  SubType: string;
  Size: string;
  UOM: string;
  Standard_Cost: string;
  MinStockLevel: string;
  Specification: string;
  Active: boolean;
  QRCodeURL: string;
}

interface IFormState {
    formData: IMaterialFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    categoryOptions: IComboBoxOption[];
    subTypeOptions: IComboBoxOption[];
    sizeOptions: IComboBoxOption[];
    uomOptions: IComboBoxOption[];
    showCustomCategory: boolean;
    showCustomUOM: boolean;
}

interface IHyperlinkFieldValue {
    Url: string;
    Description: string;
}

function getHyperlinkUrl(value: unknown): string {
    if (!value) {
        return '';
    }

    if (typeof value === 'string') {
        return value;
    }

    if (typeof value === 'object') {
        const hyperlink = value as Partial<IHyperlinkFieldValue>;
        return hyperlink.Url || hyperlink.Description || '';
    }

    return '';
}

function getHyperlinkValue(url: string): IHyperlinkFieldValue {
    return {
        Url: url,
        Description: url,
    };
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
                    SubType: '',
                    Size: '',
                    UOM: '',
                    Standard_Cost: '',
                    MinStockLevel: '',
                    Specification: '',
                    Active: true,
                    QRCodeURL: '',
                },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        categoryOptions: [],
        subTypeOptions: DEFAULT_SUB_TYPE_OPTIONS,
        sizeOptions: DEFAULT_SIZE_OPTIONS,
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
          // Use current form category/subtype/size if available, else defaults
          const cat = state.formData.Category || 'Pipes';
          const sub = state.formData.SubType || 'Carbon Steel Line';
          const sz = state.formData.Size || 'N/A';
          const catCode = getCategoryCode(cat);
          const subCode = getSubTypeCode(sub);
          const sizeCode = getSizeCode(sz);

          // Get all items and find the highest sequence for same category+subtype+size
          const allItems = await sharePointService.getListData('ENT_Materials_Master', undefined, 500, 0, undefined);
          let nextNumber = 1;
          const prefix = `${catCode}-${subCode}-${sizeCode}-`;
          if (allItems && allItems.length > 0) {
            allItems.forEach((item) => {
              if (item.Material_Code) {
                // Match same prefix pattern for sequential numbering
                if (item.Material_Code.startsWith(prefix)) {
                  const seq = parseInt(item.Material_Code.slice(prefix.length), 10);
                  if (!isNaN(seq) && seq >= nextNumber) {
                    nextNumber = seq + 1;
                  }
                }
                // Also track global max as fallback
                const matches = item.Material_Code.match(/\d{4}$/);
                if (matches) {
                  const num = parseInt(matches[0], 10);
                  if (num >= nextNumber) {
                    nextNumber = num + 1;
                  }
                }
              }
            });
          }

          const newCode = `${prefix}${nextNumber.toString().padStart(4, '0')}`;
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Material_Code: newCode },
          }));
        } catch (error) {
          console.error('Error generating material code:', error);
          // Fallback to timestamp-based code
          const cat = state.formData.Category || 'PIP';
          const sub = state.formData.SubType || 'CSL';
          const sz = state.formData.Size || '00';
          const catCode = getCategoryCode(cat);
          const subCode = getSubTypeCode(sub);
          const sizeCode = getSizeCode(sz);
          const timestamp = Date.now().toString().slice(-4);
          setState((prev) => ({
            ...prev,
            formData: { ...prev.formData, Material_Code: `${catCode}-${subCode}-${sizeCode}-${timestamp}` },
          }));
        }
      };

      const loadChoices = async (): Promise<void> => {
        try {
          const [categoryChoices, subTypeChoices, sizeChoices, uomChoices] = await Promise.all([
            sharePointService.getFieldChoices('ENT_Materials_Master', 'Category'),
            sharePointService.getFieldChoices('ENT_Materials_Master', 'SubType'),
            sharePointService.getFieldChoices('ENT_Materials_Master', 'Size'),
            sharePointService.getFieldChoices('ENT_Materials_Master', 'UOM'),
          ]);

          const categoryOptions: IComboBoxOption[] = categoryChoices.map(choice => ({
            key: choice,
            text: choice,
          }));
          categoryOptions.push({ key: 'others', text: 'Others (Add new)' });

          const subTypeOptions: IComboBoxOption[] = subTypeChoices.length > 0
            ? subTypeChoices.map(choice => ({ key: choice, text: choice }))
            : [...DEFAULT_SUB_TYPE_OPTIONS];

          if (editMaterial?.SubType && !subTypeOptions.some((option) => option.key === editMaterial.SubType)) {
            subTypeOptions.push({ key: editMaterial.SubType, text: editMaterial.SubType });
          }

          const sizeOptions: IComboBoxOption[] = sizeChoices.length > 0
            ? sizeChoices.map(choice => ({ key: choice, text: choice }))
            : [...DEFAULT_SIZE_OPTIONS];

          if (editMaterial?.Size && !sizeOptions.some((option) => option.key === editMaterial.Size)) {
            sizeOptions.push({ key: editMaterial.Size, text: editMaterial.Size });
          }

          const uomOptions: IComboBoxOption[] = uomChoices.map(choice => ({
            key: choice,
            text: choice,
          }));
          uomOptions.push({ key: 'others', text: 'Others (Add new)' });

          setState((prev) => ({
            ...prev,
            categoryOptions,
            subTypeOptions,
            sizeOptions,
            uomOptions,
          }));
        } catch (error) {
          console.error('Error loading field choices:', error);
          // Fallback to basic options
          setState((prev) => ({
            ...prev,
            categoryOptions: [{ key: 'others', text: 'Others (Add new)' }],
            subTypeOptions: DEFAULT_SUB_TYPE_OPTIONS,
            sizeOptions: DEFAULT_SIZE_OPTIONS,
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
            SubType: editMaterial.SubType || '',
            Size: editMaterial.Size || '',
            UOM: editMaterial.UOM || '',
            Standard_Cost: editMaterial.Standard_Cost?.toString() || '',
            MinStockLevel: editMaterial.MinStockLevel?.toString() || '',
            Specification: editMaterial.Specification || '',
            Active: editMaterial.Active,
            QRCodeURL: getHyperlinkUrl(editMaterial.QRCodeURL || editMaterial.qrcodeurl),
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
            SubType: '',
            Size: '',
            UOM: '',
            Standard_Cost: '',
            MinStockLevel: '',
            Specification: '',
            Active: true,
            QRCodeURL: '',
          },
        }));

        if (!state.formData.Material_Code) {
          autoGenerateCode().catch(console.error);
        }
      }
      loadChoices().catch(console.error);
    }
  }, [isOpen, editMode, editMaterial, sharePointService]);

  // Regenerate code when Category, SubType, or Size change (but only if form is open and not editing)
  React.useEffect(() => {
    if (isOpen && !editMode && state.formData.Material_Code) {
      const cat = state.formData.Category || 'Pipes';
      const sub = state.formData.SubType || 'Carbon Steel Line';
      const sz = state.formData.Size || 'N/A';
      const catCode = getCategoryCode(cat);
      const subCode = getSubTypeCode(sub);
      const sizeCode = getSizeCode(sz);
      const latestSeq = state.formData.Material_Code.match(/\d{4}$/)?.[0] || '0001';
      const prefix = `${catCode}-${subCode}-${sizeCode}-`;
      if (!state.formData.Material_Code.startsWith(prefix)) {
        setState((prev) => ({
          ...prev,
          formData: { ...prev.formData, Material_Code: `${prefix}${latestSeq}` },
        }));
      }
    }
  }, [isOpen, editMode, state.formData.Category, state.formData.SubType, state.formData.Size]);

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
      const itemData: { [key: string]: unknown } = {
        Title: state.formData.Material_Name,
        Material_Code: state.formData.Material_Code,
        Material_Name: state.formData.Material_Name,
        Category: state.formData.Category,
        SubType: state.formData.SubType,
        Size: state.formData.Size,
        UOM: state.formData.UOM,
        Standard_Cost: parseFloat(state.formData.Standard_Cost) || 0,
        MinStockLevel: parseFloat(state.formData.MinStockLevel) || 0,
        Specification: state.formData.Specification,
        Active: state.formData.Active,
      };

      if (state.formData.QRCodeURL.trim()) {
        itemData.QRCodeURL = getHyperlinkValue(state.formData.QRCodeURL);
      }

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
        const createdItem = await sharePointService.createListItem('ENT_Materials_Master', itemData);
        const qrCodeUrl = QRCodeService.generateQRCodeUrl(state.formData.Material_Code, 240);

        if (createdItem?.ID && qrCodeUrl) {
          await sharePointService.updateListItem('ENT_Materials_Master', createdItem.ID, {
            QRCodeURL: getHyperlinkValue(qrCodeUrl),
          });
        }

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
            SubType: '',
            Size: '',
            UOM: '',
            Standard_Cost: '',
            MinStockLevel: '',
            Specification: '',
            Active: true,
            QRCodeURL: '',
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
                SubType: '',
                Size: '',
                UOM: '',
                Standard_Cost: '',
                MinStockLevel: '',
            Specification: '',
            Active: true,
            QRCodeURL: '',
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
                                label="Sub Type"
                                placeholder="Select sub type"
                                allowFreeform={false}
                                autoComplete="on"
                                options={state.subTypeOptions}
                                selectedKey={state.formData.SubType || undefined}
                                onChange={(event, option, index, value) => handleComboBoxChange('SubType', option, index, value)}
                                disabled={isDisabled}
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <ComboBox
                                label="Size"
                                placeholder="Select size"
                                allowFreeform={false}
                                autoComplete="on"
                                options={state.sizeOptions}
                                selectedKey={state.formData.Size || undefined}
                                onChange={(event, option, index, value) => handleComboBoxChange('Size', option, index, value)}
                                disabled={isDisabled}
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
                            <TextField
                                label="Min Stock Level"
                                placeholder="Enter minimum stock level"
                                type="number"
                                value={state.formData.MinStockLevel}
                                onChange={(_, value) => handleInputChange('MinStockLevel', value)}
                                disabled={isDisabled}
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Specification"
                                placeholder="Enter specification details"
                                multiline
                                rows={3}
                                value={state.formData.Specification}
                                onChange={(_, value) => handleInputChange('Specification', value)}
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
