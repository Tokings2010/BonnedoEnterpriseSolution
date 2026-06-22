import * as React from 'react';
import {
  mergeStyleSets,
  Panel,
  PanelType,
  Stack,
  Text,
  Icon,
  Separator,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  ChoiceGroup,
  IChoiceGroupOption,
  DatePicker,
  DayOfWeek,
  IconButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface INewProcurementRequestPanelProps {
  isOpen: boolean;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  onDismiss: () => void;
  onRequestCreated: () => void;
}

interface ILineItem {
  id: string;
  title: string;
  materialCode: string;
  quantity: string;
  uom: string;
  specification: string;
  estimatedUnitPrice: string;
}

interface IProject {
  ID: number;
  Project_Code: string;
  Project_Name: string;
}

interface IApprovedVendor {
  Vendor_Code: string;
  Vendor_Name: string;
}

interface ISharePointListItem {
  ID: number;
  [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';

// ─── Options ─────────────────────────────────────────────────────────────────

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: 'Materials', text: 'Materials' },
  { key: 'Equipment', text: 'Equipment' },
  { key: 'Services', text: 'Services' },
  { key: 'Subcontract', text: 'Subcontract' },
];

const UOM_OPTIONS: IDropdownOption[] = [
  { key: 'EA', text: 'EA' },
  { key: 'JT', text: 'JT' },
  { key: 'MTR', text: 'MTR' },
  { key: 'KG', text: 'KG' },
  { key: 'SET', text: 'SET' },
  { key: 'LTR', text: 'LTR' },
  { key: 'ROLL', text: 'ROLL' },
  { key: 'PR', text: 'PR' },
  { key: 'BOX', text: 'BOX' },
  { key: 'PCS', text: 'PCS' },
  { key: 'LOT', text: 'LOT' },
];

const SOURCING_OPTIONS: IChoiceGroupOption[] = [
  {
    key: 'Competitive Bid',
    text: 'Competitive Bid',
    iconProps: { iconName: 'People' },
  },
  {
    key: 'Single Source',
    text: 'Single Source',
    iconProps: { iconName: 'Contact' },
  },
];

// ─── Styles ──────────────────────────────────────────────────────────────────

const getStyles = (): ReturnType<typeof mergeStyleSets> =>
  mergeStyleSets({
    panelContent: {
      padding: '0 24px 24px',
    },
    sectionTitle: {
      fontSize: '14px',
      fontWeight: '700',
      color: '#1E2532',
      marginBottom: '12px',
      marginTop: '16px',
    },
    lineItemCard: {
      backgroundColor: '#FAFBFC',
      border: '1px solid #E5E7EB',
      borderRadius: '8px',
      padding: '12px',
      marginBottom: '8px',
    },
    lineItemHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '8px',
    },
    lineItemNumber: {
      fontSize: '12px',
      fontWeight: '600',
      color: '#5A6A85',
    },
    sourcingSection: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px',
      marginTop: '8px',
    },
    justificationNote: {
      fontSize: '11px',
      color: '#DC2626',
      marginTop: '4px',
      fontStyle: 'italic',
    },
    totalEstimate: {
      backgroundColor: '#F0FDF4',
      borderRadius: '8px',
      padding: '12px 16px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginTop: '12px',
    },
  });

// ─── Component ───────────────────────────────────────────────────────────────

const NewProcurementRequestPanel: React.FC<INewProcurementRequestPanelProps> = ({
  isOpen,
  spHttpClient,
  pageContext,
  onDismiss,
  onRequestCreated,
}) => {
  // Form state
  const [title, setTitle] = React.useState('');
  const [projectCode, setProjectCode] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [requiredDate, setRequiredDate] = React.useState<Date | undefined>(undefined);
  const [budgetEstimate, setBudgetEstimate] = React.useState('');
  const [sourcingMethod, setSourcingMethod] = React.useState('Competitive Bid');
  const [assignedVendor, setAssignedVendor] = React.useState('');
  const [justification, setJustification] = React.useState('');
  const [lineItems, setLineItems] = React.useState<ILineItem[]>([
    { id: '1', title: '', materialCode: '', quantity: '', uom: 'EA', specification: '', estimatedUnitPrice: '' },
  ]);

  // Data state
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [approvedVendors, setApprovedVendors] = React.useState<IApprovedVendor[]>([]);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);

  const classNames = getStyles() as unknown as { [key: string]: string };

  // ─── Fetch Projects & Vendors ──────────────────────────────────────────────

  const fetchLookupData = React.useCallback(async (): Promise<void> => {
    try {
      // Fetch projects
      const projUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Project_Master')/items?$select=ID,Project_Code,Project_Name&$top=100`;
      const projResponse: SPHttpClientResponse = await spHttpClient.get(projUrl, SPHttpClient.configurations.v1);
      if (projResponse.ok) {
        const projData = await projResponse.json();
        setProjects((projData.value || []).map((p: ISharePointListItem) => ({
          ID: p.ID,
          Project_Code: getString(p.Project_Code),
          Project_Name: getString(p.Project_Name),
        })));
      }

      // Fetch approved vendors
      const vendorUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items?$filter=Vendor_Status eq 'Active'&$select=Vendor_Code,Vendor_Name&$top=200`;
      const vendorResponse: SPHttpClientResponse = await spHttpClient.get(vendorUrl, SPHttpClient.configurations.v1);
      if (vendorResponse.ok) {
        const vendorData = await vendorResponse.json();
        setApprovedVendors((vendorData.value || []).map((v: ISharePointListItem) => ({
          Vendor_Code: getString(v.Vendor_Code),
          Vendor_Name: getString(v.Vendor_Name),
        })));
      }
    } catch (err) {
      console.error('Error fetching lookup data:', err);
    }
  }, [spHttpClient, pageContext]);

  // ─── Form Helpers ──────────────────────────────────────────────────────────

  const resetForm = (): void => {
    setTitle('');
    setProjectCode('');
    setCategory('');
    setRequiredDate(undefined);
    setBudgetEstimate('');
    setSourcingMethod('Competitive Bid');
    setAssignedVendor('');
    setJustification('');
    setLineItems([
      { id: '1', title: '', materialCode: '', quantity: '', uom: 'EA', specification: '', estimatedUnitPrice: '' },
    ]);
    setMessage(undefined);
  };

  React.useEffect(() => {
    if (isOpen) {
      fetchLookupData().catch(() => undefined);
      resetForm();
    }
  }, [isOpen, fetchLookupData, resetForm]);

  const addLineItem = (): void => {
    const newId = (lineItems.length + 1).toString();
    setLineItems([
      ...lineItems,
      { id: newId, title: '', materialCode: '', quantity: '', uom: 'EA', specification: '', estimatedUnitPrice: '' },
    ]);
  };

  const removeLineItem = (id: string): void => {
    if (lineItems.length <= 1) return;
    setLineItems(lineItems.filter((item) => item.id !== id));
  };

  const updateLineItem = (id: string, field: keyof ILineItem, value: string): void => {
    setLineItems(lineItems.map((item) =>
      item.id === id ? { ...item, [field]: value } : item
    ));
  };

  // ─── Calculate Total ───────────────────────────────────────────────────────

  const totalEstimate = lineItems.reduce((sum, item) => {
    const qty = parseFloat(item.quantity) || 0;
    const price = parseFloat(item.estimatedUnitPrice) || 0;
    return sum + qty * price;
  }, 0);

  // ─── Submit ────────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    // Validation
    if (!title.trim()) {
      setMessage({ type: MessageBarType.warning, text: 'Please enter a title/description' });
      return;
    }
    if (!category) {
      setMessage({ type: MessageBarType.warning, text: 'Please select a category' });
      return;
    }
    if (!requiredDate) {
      setMessage({ type: MessageBarType.warning, text: 'Please select a required-by date' });
      return;
    }
    if (sourcingMethod === 'Single Source' && !justification.trim()) {
      setMessage({ type: MessageBarType.warning, text: 'Justification is required for Single Source procurement' });
      return;
    }
    if (sourcingMethod === 'Single Source' && !assignedVendor) {
      setMessage({ type: MessageBarType.warning, text: 'Please select a vendor for Single Source' });
      return;
    }

    setIsSubmitting(true);
    setMessage(undefined);

    try {
      // 1. Generate RequestID (PR-YYYY-NNNN)
      const year = new Date().getFullYear().toString();
      const prefix = `PR-${year}-`;
      const idQueryUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Procurement_Requests')/items?$filter=startswith(RequestID,'${prefix}')&$orderby=RequestID desc&$top=1&$select=RequestID`;
      let nextNumber = 1;
      try {
        const idResponse: SPHttpClientResponse = await spHttpClient.get(
          idQueryUrl,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (idResponse.ok) {
          const idData = await idResponse.json();
          if (idData.value && idData.value.length > 0) {
            const lastId = idData.value[0].RequestID || '';
            const lastNum = parseInt(lastId.split('-').pop() || '0', 10);
            if (!isNaN(lastNum)) nextNumber = lastNum + 1;
          }
        }
      } catch { /* first item — start at 1 */ }

      const generatedRequestID = `${prefix}${nextNumber.toString().padStart(4, '0')}`;

      // 2. Create the procurement request
      const requestUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Procurement_Requests')/items`;
      const requestBody: Record<string, unknown> = {
        Title: title,
        RequestID: generatedRequestID,
        Project_Code: projectCode || undefined,
        Request_Date: new Date().toISOString(),
        Required_Date: requiredDate.toISOString(),
        Sourcing_Method: sourcingMethod,
        Status: 'Draft',
        Category: category,
      };

      if (budgetEstimate) requestBody.Budget_Estimate = parseFloat(budgetEstimate);
      if (sourcingMethod === 'Single Source') {
        requestBody.Assigned_Vendor = assignedVendor;
        requestBody.Justification = justification;
      }

      const requestResponse: SPHttpClientResponse = await spHttpClient.post(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
          body: JSON.stringify(requestBody),
        }
      );

      if (!requestResponse.ok) {
        throw new Error(`Failed to create procurement request: ${requestResponse.status}`);
      }

      const requestID = generatedRequestID;

      // 3. Create line items
      const validLineItems = lineItems.filter((li) => li.title.trim());
      for (const lineItem of validLineItems) {
        const itemUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Procurement_Request_Items')/items`;
        const itemBody: Record<string, unknown> = {
          Title: lineItem.title,
          RequestID: requestID,
          Quantity: parseFloat(lineItem.quantity) || 0,
          UOM: lineItem.uom,
        };

        if (lineItem.materialCode) itemBody.Material_Code = lineItem.materialCode;
        if (lineItem.specification) itemBody.Specification = lineItem.specification;
        if (lineItem.estimatedUnitPrice) itemBody.Estimated_Unit_Price = parseFloat(lineItem.estimatedUnitPrice);

        await spHttpClient.post(
          itemUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
            },
            body: JSON.stringify(itemBody),
          }
        );
      }

      setMessage({ type: MessageBarType.success, text: 'Procurement request created successfully!' });
      setTimeout(() => onRequestCreated(), 1500);
    } catch (err) {
      console.error('Error creating procurement request:', err);
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to create request' });
    } finally {
      setIsSubmitting(false);
    }
  };

  // ─── Dropdown Options ──────────────────────────────────────────────────────

  const projectOptions: IDropdownOption[] = [
    { key: '', text: '— No Project —' },
    ...projects.map((p) => ({ key: p.Project_Code, text: `${p.Project_Code} — ${p.Project_Name}` })),
  ];

  const vendorOptions: IDropdownOption[] = approvedVendors.map((v) => ({
    key: v.Vendor_Code,
    text: `${v.Vendor_Name} (${v.Vendor_Code})`,
  }));

  // ─── Render ────────────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="New Procurement Request"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 24px' }}>
          <PrimaryButton
            text="Create Request"
            onClick={handleSubmit}
            disabled={isSubmitting}
            iconProps={{ iconName: 'CheckMark' }}
          />
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSubmitting} />
        </Stack>
      )}
      isFooterAtBottom={true}
    >
      <div className={classNames.panelContent}>
        {/* Message */}
        {message && (
          <MessageBar
            messageBarType={message.type}
            isMultiline={false}
            onDismiss={() => setMessage(undefined)}
            style={{ marginBottom: '12px' }}
          >
            {message.text}
          </MessageBar>
        )}

        {/* Basic Info */}
        <div className={classNames.sectionTitle}>Request Details</div>
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label="Title / Description"
            value={title}
            onChange={(_, v) => setTitle(v || '')}
            required
            placeholder="Brief description of what you need to procure"
          />
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <Dropdown
              label="Project Code"
              placeholder="Select project..."
              selectedKey={projectCode}
              options={projectOptions}
              onChange={(_, option) => setProjectCode(option?.key as string || '')}
              styles={{ root: { flex: 1 } }}
            />
            <Dropdown
              label="Category"
              placeholder="Select..."
              selectedKey={category}
              options={CATEGORY_OPTIONS}
              onChange={(_, option) => setCategory(option?.key as string || '')}
              required
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <DatePicker
              label="Required By"
              value={requiredDate}
              onSelectDate={(date) => setRequiredDate(date || undefined)}
              firstDayOfWeek={DayOfWeek.Monday}
              placeholder="Select date..."
              isRequired
              styles={{ root: { flex: 1 } }}
            />
            <TextField
              label="Budget Estimate (₦)"
              value={budgetEstimate}
              onChange={(_, v) => setBudgetEstimate(v || '')}
              type="number"
              placeholder="0.00"
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        </Stack>

        <Separator />

        {/* Sourcing Method */}
        <div className={classNames.sectionTitle}>
          <Icon iconName="Switch" style={{ marginRight: '6px' }} />
          Sourcing Method
        </div>
        <div className={classNames.sourcingSection}>
          <ChoiceGroup
            selectedKey={sourcingMethod}
            options={SOURCING_OPTIONS}
            onChange={(_, option) => setSourcingMethod(option?.key || 'Competitive Bid')}
            styles={{
              flexContainer: { display: 'flex', gap: '24px' },
            }}
          />

          {sourcingMethod === 'Single Source' && (
            <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: '12px' }}>
              <Dropdown
                label="Assigned Vendor"
                placeholder="Select approved vendor..."
                selectedKey={assignedVendor}
                options={vendorOptions}
                onChange={(_, option) => setAssignedVendor(option?.key as string || '')}
                required
              />
              <TextField
                label="Justification"
                value={justification}
                onChange={(_, v) => setJustification(v || '')}
                multiline
                rows={3}
                required
                placeholder="Explain why this must be single-sourced (e.g., OEM-specific, emergency, sole supplier)..."
              />
              <span className={classNames.justificationNote}>
                * Justification is mandatory for single-source procurement
              </span>
            </Stack>
          )}

          {sourcingMethod === 'Competitive Bid' && (
            <Text variant="small" style={{ display: 'block', marginTop: '8px', color: '#5A6A85' }}>
              After creating this request, you can invite approved vendors to submit quotes from the Bid Detail view.
            </Text>
          )}
        </div>

        <Separator />

        {/* Line Items */}
        <div className={classNames.sectionTitle}>
          <Icon iconName="BulletedList" style={{ marginRight: '6px' }} />
          Line Items
        </div>

        {lineItems.map((item, index) => (
          <div key={item.id} className={classNames.lineItemCard}>
            <div className={classNames.lineItemHeader}>
              <span className={classNames.lineItemNumber}>Item #{index + 1}</span>
              {lineItems.length > 1 && (
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Remove item"
                  styles={{ root: { color: '#EF4444', height: '24px', width: '24px' } }}
                  onClick={() => removeLineItem(item.id)}
                />
              )}
            </div>
            <Stack tokens={{ childrenGap: 8 }}>
              <TextField
                placeholder="Item description"
                value={item.title}
                onChange={(_, v) => updateLineItem(item.id, 'title', v || '')}
              />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField
                  placeholder="Material code (optional)"
                  value={item.materialCode}
                  onChange={(_, v) => updateLineItem(item.id, 'materialCode', v || '')}
                  styles={{ root: { flex: 2 } }}
                />
                <TextField
                  placeholder="Qty"
                  value={item.quantity}
                  onChange={(_, v) => updateLineItem(item.id, 'quantity', v || '')}
                  type="number"
                  styles={{ root: { flex: 1 } }}
                />
                <Dropdown
                  placeholder="UOM"
                  selectedKey={item.uom}
                  options={UOM_OPTIONS}
                  onChange={(_, option) => updateLineItem(item.id, 'uom', option?.key as string || 'EA')}
                  styles={{ root: { flex: 1 } }}
                />
                <TextField
                  placeholder="Unit price"
                  value={item.estimatedUnitPrice}
                  onChange={(_, v) => updateLineItem(item.id, 'estimatedUnitPrice', v || '')}
                  type="number"
                  styles={{ root: { flex: 1 } }}
                />
              </Stack>
              <TextField
                placeholder="Technical specification (optional)"
                value={item.specification}
                onChange={(_, v) => updateLineItem(item.id, 'specification', v || '')}
              />
            </Stack>
          </div>
        ))}

        <DefaultButton
          text="Add Line Item"
          iconProps={{ iconName: 'Add' }}
          onClick={addLineItem}
          styles={{ root: { marginTop: '8px' } }}
        />

        {/* Total Estimate */}
        {totalEstimate > 0 && (
          <div className={classNames.totalEstimate}>
            <Text style={{ fontWeight: 600, color: '#065F46' }}>Estimated Total</Text>
            <Text style={{ fontWeight: 700, fontSize: '16px', color: '#065F46' }}>
              ₦{totalEstimate.toLocaleString()}
            </Text>
          </div>
        )}

        {isSubmitting && (
          <Stack horizontalAlign="center" style={{ marginTop: '16px' }}>
            <Spinner size={SpinnerSize.small} label="Creating request..." />
          </Stack>
        )}
      </div>
    </Panel>
  );
};

export default NewProcurementRequestPanel;
