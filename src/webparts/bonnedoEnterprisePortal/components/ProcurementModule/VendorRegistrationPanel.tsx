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

export interface IVendorRegistrationPanelProps {
    isOpen: boolean;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onDismiss: () => void;
    onVendorRegistered: () => void;
}

interface IComplianceDocument {
    id: string;
    documentType: string;
    status: string;
    remarks: string;
    expiryDate: Date | undefined;
    file: File | undefined;
}

interface ISharePointListItem {
    ID: number;
    [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';
const getOptionalString = (value: unknown): string | undefined => typeof value === 'string' && value ? value : undefined;

const CATEGORY_OPTIONS: IDropdownOption[] = [
    { key: 'Materials', text: 'Materials' },
    { key: 'Equipment', text: 'Equipment' },
    { key: 'Services', text: 'Services' },
    { key: 'Subcontract', text: 'Subcontract' },
    { key: 'Logistics', text: 'Logistics' },
    { key: 'IT Services', text: 'IT Services' },
    { key: 'Other', text: 'Other' },
];

const REGISTRATION_TYPE_OPTIONS: IDropdownOption[] = [
    { key: 'New Vendor', text: 'New Vendor' },
    { key: 'Existing Vendor Renewal', text: 'Existing Vendor Renewal' },
    { key: 'Reactivated Vendor', text: 'Reactivated Vendor' },
];

const DOCUMENT_TYPE_OPTIONS: IDropdownOption[] = [
    { key: 'Certificate of Incorporation', text: 'Certificate of Incorporation' },
    { key: 'Tax Clearance Certificate', text: 'Tax Clearance Certificate' },
    { key: 'VAT Registration', text: 'VAT Registration' },
    { key: 'Insurance Certificate', text: 'Insurance Certificate' },
    { key: 'HSE Policy/Certification', text: 'HSE Policy/Certification' },
    { key: 'Bank Details/Reference Letter', text: 'Bank Details/Reference Letter' },
    { key: 'Company Profile', text: 'Company Profile' },
    { key: 'Previous Work Experience', text: 'Previous Work Experience' },
    { key: 'NIPEX/DPR/NUPRC Registration', text: 'NIPEX/DPR/NUPRC Registration' },
];

const DOCUMENT_STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'Valid', text: 'Valid' },
    { key: 'Under Review', text: 'Under Review' },
    { key: 'Expired', text: 'Expired' },
];

const REGISTRATION_TYPE_CHOICES: IChoiceGroupOption[] = [
    { key: 'New Vendor', text: 'New Vendor', iconProps: { iconName: 'AddFriend' } },
    { key: 'Existing Vendor Renewal', text: 'Existing Vendor Renewal', iconProps: { iconName: 'Refresh' } },
    { key: 'Reactivated Vendor', text: 'Reactivated Vendor', iconProps: { iconName: 'Redo' } },
];

const getStyles = (): ReturnType<typeof mergeStyleSets> =>
    mergeStyleSets({
        panelContent: {
            padding: '0 24px 24px',
        },
        sectionTitle: {
            fontSize: '14px',
            fontWeight: 700,
            color: '#1E2532',
            marginBottom: 12,
            marginTop: 16,
        },
        documentCard: {
            backgroundColor: '#FAFBFC',
            border: '1px solid #E5E7EB',
            borderRadius: 8,
            padding: 12,
            marginBottom: 8,
        },
        documentHeader: {
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: 8,
        },
        documentNumber: {
            fontSize: 12,
            fontWeight: 600,
            color: '#5A6A85',
        },
        uploadBox: {
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            flexWrap: 'wrap',
            marginTop: 8,
        },
        uploadHint: {
            fontSize: 12,
            color: '#5A6A85',
        },
        summaryCard: {
            backgroundColor: '#F0FDF4',
            borderRadius: 8,
            padding: 12,
            marginTop: 12,
        },
    });

const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> =>
    new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (): void => {
            resolve(reader.result as ArrayBuffer);
        };

        reader.onerror = (): void => {
            reject(reader.error || new Error('Failed to read file'));
        };

        reader.readAsArrayBuffer(file);
    });

const VendorRegistrationPanel: React.FC<IVendorRegistrationPanelProps> = ({
    isOpen,
    spHttpClient,
    pageContext,
    onDismiss,
    onVendorRegistered,
}) => {
    const [vendorCode, setVendorCode] = React.useState('');
    const [vendorName, setVendorName] = React.useState('');
    const [contactPerson, setContactPerson] = React.useState('');
    const [email, setEmail] = React.useState('');
    const [phone, setPhone] = React.useState('');
    const [address, setAddress] = React.useState('');
    const [registrationType, setRegistrationType] = React.useState('New Vendor');
    const [vendorCategory, setVendorCategory] = React.useState('');
    const [taxId, setTaxId] = React.useState('');
    const [rcNumber, setRcNumber] = React.useState('');
    const [bankName, setBankName] = React.useState('');
    const [bankAccount, setBankAccount] = React.useState('');
    const [notes, setNotes] = React.useState('');
    const [documents, setDocuments] = React.useState<IComplianceDocument[]>([
        { id: '1', documentType: 'Certificate of Incorporation', status: 'Under Review', remarks: '', expiryDate: undefined, file: undefined },
    ]);
    const [isSubmitting, setIsSubmitting] = React.useState(false);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
    const fileInputRef = React.useRef<HTMLInputElement | null>(null);
    const classNames = getStyles() as unknown as { [key: string]: string };

    React.useEffect(() => {
        if (isOpen) {
            resetForm().catch(() => undefined);
        }
    }, [isOpen]);

    const resetForm = React.useCallback(async (): Promise<void> => {
        const prefix = 'VND-';
        const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items?$filter=startswith(Vendor_Code,'${prefix}')&$orderby=Vendor_Code desc&$top=1&$select=Vendor_Code`;
        let nextNumber = 1;

        try {
            const response: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: { Accept: 'application/json;odata=nometadata' },
            });

            if (response.ok) {
                const data = await response.json();
                const items: ISharePointListItem[] = data.value || [];
                const lastCode = getString(items[0]?.Vendor_Code);
                const lastNumber = parseInt(lastCode.replace(prefix, ''), 10);

                if (!isNaN(lastNumber)) {
                    nextNumber = lastNumber + 1;
                }
            }
        } catch (error) {
            console.error('Error generating vendor code:', error);
        }

        setVendorCode(`${prefix}${nextNumber.toString().padStart(4, '0')}`);
        setVendorName('');
        setContactPerson('');
        setEmail('');
        setPhone('');
        setAddress('');
        setRegistrationType('New Vendor');
        setVendorCategory('');
        setTaxId('');
        setRcNumber('');
        setBankName('');
        setBankAccount('');
        setNotes('');
        setDocuments([
            { id: '1', documentType: 'Certificate of Incorporation', status: 'Under Review', remarks: '', expiryDate: undefined, file: undefined },
        ]);
        setMessage(undefined);
    }, [pageContext, spHttpClient]);

    const addDocument = (): void => {
        const newId = (documents.length + 1).toString();
        setDocuments([
            ...documents,
            { id: newId, documentType: 'Tax Clearance Certificate', status: 'Under Review', remarks: '', expiryDate: undefined, file: undefined },
        ]);
    };

    const removeDocument = (id: string): void => {
        if (documents.length <= 1) {
            return;
        }

        setDocuments(documents.filter((document) => document.id !== id));
    };

    const updateDocument = (id: string, field: keyof IComplianceDocument, value: string | Date | File | undefined): void => {
        setDocuments(documents.map((document) =>
            document.id === id ? { ...document, [field]: value } : document
        ));
    };

    const validateForm = (): boolean => {
        if (!vendorName.trim()) {
            setMessage({ type: MessageBarType.warning, text: 'Vendor name is required' });
            return false;
        }

        if (!contactPerson.trim()) {
            setMessage({ type: MessageBarType.warning, text: 'Contact person is required' });
            return false;
        }

        if (!email.trim()) {
            setMessage({ type: MessageBarType.warning, text: 'Email address is required' });
            return false;
        }

        if (!vendorCategory) {
            setMessage({ type: MessageBarType.warning, text: 'Vendor category is required' });
            return false;
        }

        return true;
    };

    const uploadDocument = async (document: IComplianceDocument, createdVendorCode: string): Promise<void> => {
        if (!document.file) {
            return;
        }

        const libraryServerRelativeUrl = `${pageContext.web.serverRelativeUrl}/Vendor_Documents_Library`;
        const uploadUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryServerRelativeUrl}')/Files/add(url='${encodeURIComponent(document.file.name)}',overwrite=true)`;
        const fileBuffer = await readFileAsArrayBuffer(document.file);

        const uploadResponse: SPHttpClientResponse = await spHttpClient.post(
            uploadUrl,
            SPHttpClient.configurations.v1,
            {
                headers: { Accept: 'application/json;odata=nometadata' },
                body: fileBuffer,
            }
        );

        if (!uploadResponse.ok) {
            throw new Error(`Failed to upload ${document.file.name}: ${uploadResponse.status}`);
        }

        const uploadData = await uploadResponse.json();
        const serverRelativeUrl = getString(uploadData.ServerRelativeUrl || uploadData.ServerRelativePath);

        if (!serverRelativeUrl) {
            throw new Error(`Failed to resolve uploaded file path for ${document.file.name}`);
        }

        const listItemUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(serverRelativeUrl)}')/ListItemAllFields`;
        const documentUrl = `${pageContext.web.absoluteUrl}${serverRelativeUrl}`;
        const metadataBody: Record<string, unknown> = {
            Title: document.documentType,
            VendorID: createdVendorCode,
            Document_Type: document.documentType,
            Document_URL: {
                Url: documentUrl,
                Description: document.file.name,
            },
            Status: document.status,
            Remarks: document.remarks,
        };

        if (document.expiryDate) {
            metadataBody.Expiry_Date = document.expiryDate.toISOString();
        }

        const metadataResponse: SPHttpClientResponse = await spHttpClient.post(
            listItemUrl,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    Accept: 'application/json;odata=nometadata',
                    'Content-Type': 'application/json;odata=nometadata',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                },
                body: JSON.stringify(metadataBody),
            }
        );

        if (!metadataResponse.ok) {
            throw new Error(`Failed to store document metadata for ${document.file.name}: ${metadataResponse.status}`);
        }
    };

    const handleSubmit = async (): Promise<void> => {
        if (!validateForm()) {
            return;
        }

        setIsSubmitting(true);
        setMessage(undefined);

        try {
            const requestBody: Record<string, unknown> = {
                Title: vendorName,
                Vendor_Code: vendorCode,
                Vendor_Name: vendorName,
                Contact_Person: contactPerson,
                Email: email,
                Phone: phone,
                Address: address,
                Registration_Type: registrationType,
                Vendor_Category: vendorCategory,
                Tax_ID: taxId,
                RC_Number: rcNumber,
                Bank_Name: bankName,
                Bank_Account: bankAccount,
                Vendor_Status: 'Pending',
                Notes: notes,
            };

            const vendorResponse: SPHttpClientResponse = await spHttpClient.post(
                `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        Accept: 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata',
                    },
                    body: JSON.stringify(requestBody),
                }
            );

            if (!vendorResponse.ok) {
                throw new Error(`Failed to register vendor: ${vendorResponse.status}`);
            }

            const validDocuments = documents.filter((document) => document.file);
            for (const document of validDocuments) {
                await uploadDocument(document, vendorCode);
            }

            setMessage({ type: MessageBarType.success, text: 'Vendor registered successfully' });
            setTimeout(() => {
                onVendorRegistered();
            }, 1200);
        } catch (error) {
            console.error('Error registering vendor:', error);
            setMessage({ type: MessageBarType.error, text: error instanceof Error ? error.message : 'Failed to register vendor' });
        } finally {
            setIsSubmitting(false);
        }
    };

    const totalDocuments = documents.filter((document) => document.file).length;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.large}
            headerText="Register Vendor"
            closeButtonAriaLabel="Close"
            isFooterAtBottom={true}
            onRenderFooterContent={() => (
                <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 24px' }}>
                    <PrimaryButton
                        text="Register Vendor"
                        onClick={handleSubmit}
                        disabled={isSubmitting}
                        iconProps={{ iconName: 'CheckMark' }}
                    />
                    <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSubmitting} />
                </Stack>
            )}
        >
            <div className={classNames.panelContent}>
                {message && (
                    <MessageBar
                        messageBarType={message.type}
                        isMultiline={false}
                        onDismiss={() => setMessage(undefined)}
                        style={{ marginBottom: 12 }}
                    >
                        {message.text}
                    </MessageBar>
                )}

                <div className={classNames.sectionTitle}>Vendor Details</div>
                <Stack tokens={{ childrenGap: 12 }}>
                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Vendor Code"
                            value={vendorCode}
                            disabled={true}
                            styles={{ root: { flex: 1 } }}
                        />
                        <Dropdown
                            label="Registration Type"
                            selectedKey={registrationType}
                            options={REGISTRATION_TYPE_OPTIONS}
                            onChange={(_, option) => setRegistrationType(option?.key as string || 'New Vendor')}
                            styles={{ root: { flex: 1 } }}
                        />
                    </Stack>

                    <TextField
                        label="Vendor Name"
                        value={vendorName}
                        onChange={(_, value) => setVendorName(value || '')}
                        required
                        placeholder="Enter legal vendor name"
                    />

                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Contact Person"
                            value={contactPerson}
                            onChange={(_, value) => setContactPerson(value || '')}
                            required
                            styles={{ root: { flex: 1 } }}
                        />
                        <TextField
                            label="Email"
                            value={email}
                            onChange={(_, value) => setEmail(value || '')}
                            required
                            styles={{ root: { flex: 1 } }}
                        />
                    </Stack>

                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Phone"
                            value={phone}
                            onChange={(_, value) => setPhone(value || '')}
                            styles={{ root: { flex: 1 } }}
                        />
                        <Dropdown
                            label="Vendor Category"
                            placeholder="Select category"
                            selectedKey={vendorCategory}
                            options={CATEGORY_OPTIONS}
                            onChange={(_, option) => setVendorCategory(option?.key as string || '')}
                            required
                            styles={{ root: { flex: 1 } }}
                        />
                    </Stack>

                    <TextField
                        label="Address"
                        value={address}
                        onChange={(_, value) => setAddress(value || '')}
                        multiline
                        rows={3}
                    />

                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Tax ID (TIN)"
                            value={taxId}
                            onChange={(_, value) => setTaxId(value || '')}
                            styles={{ root: { flex: 1 } }}
                        />
                        <TextField
                            label="RC Number (CAC)"
                            value={rcNumber}
                            onChange={(_, value) => setRcNumber(value || '')}
                            styles={{ root: { flex: 1 } }}
                        />
                    </Stack>

                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Bank Name"
                            value={bankName}
                            onChange={(_, value) => setBankName(value || '')}
                            styles={{ root: { flex: 1 } }}
                        />
                        <TextField
                            label="Bank Account"
                            value={bankAccount}
                            onChange={(_, value) => setBankAccount(value || '')}
                            styles={{ root: { flex: 1 } }}
                        />
                    </Stack>

                    <TextField
                        label="Notes"
                        value={notes}
                        onChange={(_, value) => setNotes(value || '')}
                        multiline
                        rows={3}
                    />
                </Stack>

                <Separator />

                <div className={classNames.sectionTitle}>
                    <Icon iconName="ContactCard" style={{ marginRight: 6 }} />
                    Registration Type
                </div>
                <ChoiceGroup
                    selectedKey={registrationType}
                    options={REGISTRATION_TYPE_CHOICES}
                    onChange={(_, option) => setRegistrationType(option?.key || 'New Vendor')}
                    styles={{ flexContainer: { display: 'flex', gap: 24 } }}
                />

                <Separator />

                <div className={classNames.sectionTitle}>
                    <Icon iconName="Documentation" style={{ marginRight: 6 }} />
                    Compliance Documents
                </div>

                {documents.map((document, index) => (
                    <div key={document.id} className={classNames.documentCard}>
                        <div className={classNames.documentHeader}>
                            <span className={classNames.documentNumber}>Document {index + 1}</span>
                            {documents.length > 1 && (
                                <IconButton
                                    iconProps={{ iconName: 'Delete' }}
                                    title="Remove document"
                                    styles={{ root: { color: '#EF4444', height: 24, width: 24 } }}
                                    onClick={() => removeDocument(document.id)}
                                />
                            )}
                        </div>

                        <Stack tokens={{ childrenGap: 8 }}>
                            <Stack horizontal tokens={{ childrenGap: 8 }}>
                                <Dropdown
                                    label="Document Type"
                                    selectedKey={document.documentType}
                                    options={DOCUMENT_TYPE_OPTIONS}
                                    onChange={(_, option) => updateDocument(document.id, 'documentType', option?.key as string || '')}
                                    styles={{ root: { flex: 2 } }}
                                />
                                <Dropdown
                                    label="Status"
                                    selectedKey={document.status}
                                    options={DOCUMENT_STATUS_OPTIONS}
                                    onChange={(_, option) => updateDocument(document.id, 'status', option?.key as string || 'Under Review')}
                                    styles={{ root: { flex: 1 } }}
                                />
                            </Stack>

                            <Stack horizontal tokens={{ childrenGap: 8 }}>
                                <DatePicker
                                    label="Expiry Date"
                                    value={document.expiryDate}
                                    onSelectDate={(date) => updateDocument(document.id, 'expiryDate', date || undefined)}
                                    firstDayOfWeek={DayOfWeek.Monday}
                                    placeholder="Select expiry date..."
                                    styles={{ root: { flex: 1 } }}
                                />
                                <div className={classNames.uploadBox}>
                                    <input
                                        ref={fileInputRef}
                                        type="file"
                                        style={{ display: 'none' }}
                                        onChange={(event) => {
                                            const selectedFile = event.currentTarget.files?.[0];
                                            if (selectedFile) {
                                                updateDocument(document.id, 'file', selectedFile);
                                            }
                                            event.currentTarget.value = '';
                                        }}
                                    />
                                    <DefaultButton
                                        text={document.file ? document.file.name : 'Choose File'}
                                        iconProps={{ iconName: document.file ? 'Attach' : 'Upload' }}
                                        onClick={() => fileInputRef.current?.click()}
                                    />
                                    {document.file && <span className={classNames.uploadHint}>{(document.file.size / 1024).toFixed(0)} KB</span>}
                                </div>
                            </Stack>

                            <TextField
                                label="Remarks"
                                value={document.remarks}
                                onChange={(_, value) => updateDocument(document.id, 'remarks', value || '')}
                                multiline
                                rows={2}
                            />
                        </Stack>
                    </div>
                ))}

                <DefaultButton
                    text="Add Document"
                    iconProps={{ iconName: 'Add' }}
                    onClick={addDocument}
                    styles={{ root: { marginTop: 8 } }}
                />

                {totalDocuments > 0 && (
                    <div className={classNames.summaryCard}>
                        <Text variant="small" style={{ fontWeight: 600, color: '#065F46' }}>
                            {totalDocuments} compliance document(s) will be uploaded to Vendor_Documents_Library.
                        </Text>
                    </div>
                )}

                {isSubmitting && (
                    <Stack horizontalAlign="center" style={{ marginTop: 16 }}>
                        <Spinner size={SpinnerSize.small} label="Registering vendor and uploading documents..." />
                    </Stack>
                )}
            </div>
        </Panel>
    );
};

export default VendorRegistrationPanel;
