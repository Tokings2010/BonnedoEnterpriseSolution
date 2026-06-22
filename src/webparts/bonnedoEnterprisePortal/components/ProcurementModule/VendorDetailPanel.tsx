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
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IVendorItem } from './VendorRegistrationTab';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IVendorDetailPanelProps {
  isOpen: boolean;
  vendor: IVendorItem | undefined;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  onDismiss: () => void;
  onVendorUpdated: () => void;
}

interface IVendorDocument {
  ID: number;
  Title: string;
  VendorID: string;
  Document_Type: string;
  Document_URL: string | undefined;
  Expiry_Date: string | undefined;
  Upload_Date: string | undefined;
  Status: string;
  Remarks: string;
}

// ─── Required Document Types ─────────────────────────────────────────────────

const REQUIRED_DOCUMENTS: string[] = [
  'Certificate of Incorporation',
  'Tax Clearance Certificate',
  'VAT Registration',
  'Insurance Certificate',
  'HSE Policy/Certification',
  'Bank Details/Reference Letter',
  'Company Profile',
  'Previous Work Experience',
  'NIPEX/DPR/NUPRC Registration',
  'Other',
];

const DOC_STATUS_COLORS: Record<string, { bg: string; text: string; icon: string }> = {
  Valid: { bg: '#D1FAE5', text: '#065F46', icon: 'CheckMark' },
  Expired: { bg: '#FEE2E2', text: '#991B1B', icon: 'Warning' },
  'Under Review': { bg: '#DBEAFE', text: '#1E40AF', icon: 'Clock' },
  Rejected: { bg: '#FEE2E2', text: '#991B1B', icon: 'Cancel' },
  Missing: { bg: '#F3F4F6', text: '#6B7280', icon: 'StatusCircleQuestionMark' },
};

// ─── Styles ──────────────────────────────────────────────────────────────────

interface ISharePointListItem {
  ID: number;
  [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';
const getOptionalString = (value: unknown): string | undefined => typeof value === 'string' && value ? value : undefined;
const getHyperlinkUrl = (value: unknown): string | undefined => {
  if (typeof value === 'string') {
    return value;
  }

  if (value && typeof value === 'object' && 'Url' in value) {
    const url = (value as { Url?: unknown }).Url;
    return typeof url === 'string' ? url : undefined;
  }

  return undefined;
};

const getStyles = (): ReturnType<typeof mergeStyleSets> =>
  mergeStyleSets({
    panelContent: {
      padding: '0 24px 24px',
    },
    vendorHeader: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px',
      marginBottom: '16px',
    },
    vendorName: {
      fontSize: '18px',
      fontWeight: '700',
      color: '#1E2532',
      marginBottom: '4px',
    },
    vendorId: {
      fontSize: '13px',
      fontFamily: 'Cascadia Code, Consolas, monospace',
      color: '#5A6A85',
    },
    infoGrid: {
      display: 'grid',
      gridTemplateColumns: '1fr 1fr',
      gap: '12px',
      marginBottom: '16px',
    },
    infoItem: {
      display: 'flex',
      flexDirection: 'column',
      gap: '2px',
    },
    infoLabel: {
      fontSize: '11px',
      fontWeight: '600',
      color: '#5A6A85',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
    },
    infoValue: {
      fontSize: '14px',
      color: '#1E2532',
    },
    sectionTitle: {
      fontSize: '14px',
      fontWeight: '700',
      color: '#1E2532',
      marginBottom: '12px',
      marginTop: '8px',
    },
    docRow: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      padding: '8px 12px',
      borderRadius: '6px',
      marginBottom: '6px',
      backgroundColor: '#FAFBFC',
      border: '1px solid #E5E7EB',
    },
    docName: {
      fontSize: '13px',
      color: '#1E2532',
      flex: 1,
    },
    docStatus: {
      display: 'inline-flex',
      alignItems: 'center',
      gap: '4px',
      padding: '2px 8px',
      borderRadius: '10px',
      fontSize: '11px',
      fontWeight: '600',
    },
    statusPill: {
      display: 'inline-block',
      padding: '4px 12px',
      borderRadius: '12px',
      fontSize: '12px',
      fontWeight: '600',
    },
    actionButtons: {
      display: 'flex',
      gap: '8px',
      marginTop: '16px',
    },
  });

// ─── Component ───────────────────────────────────────────────────────────────

const VendorDetailPanel: React.FC<IVendorDetailPanelProps> = ({
  isOpen,
  vendor,
  spHttpClient,
  pageContext,
  onDismiss,
  onVendorUpdated,
}) => {
  const [documents, setDocuments] = React.useState<IVendorDocument[]>([]);
  const [isLoadingDocs, setIsLoadingDocs] = React.useState(false);
  const [isUpdating, setIsUpdating] = React.useState(false);
  const [updateMessage, setUpdateMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);

  const classNames = getStyles() as unknown as { [key: string]: string };

  // ─── Fetch Vendor Documents ────────────────────────────────────────────────

  const fetchDocuments = React.useCallback(async (): Promise<void> => {
    if (!vendor) return;
    setIsLoadingDocs(true);

    try {
      const listUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Vendor_Documents')/items?$filter=VendorID eq '${vendor.Vendor_Code}'&$orderby=Document_Type`;
      const listResponse: SPHttpClientResponse = await spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      if (!listResponse.ok) {
        throw new Error(`Failed to fetch documents: ${listResponse.status}`);
      }

      const listData = await listResponse.json();
      const listDocs: IVendorDocument[] = (listData.value || []).map((item: ISharePointListItem) => ({
        ID: item.ID,
        Title: getString(item.Title),
        VendorID: getString(item.VendorID),
        Document_Type: getString(item.Document_Type),
        Document_URL: getHyperlinkUrl(item.Document_URL),
        Expiry_Date: getOptionalString(item.Expiry_Date),
        Upload_Date: getOptionalString(item.Upload_Date),
        Status: getString(item.Status) || 'Under Review',
        Remarks: getString(item.Remarks),
      }));

      const libraryDocs: IVendorDocument[] = [];
      try {
        const libraryUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Vendor_Documents_Library')/items?$filter=VendorID eq '${vendor.Vendor_Code}'&$select=ID,Title,FileRef,FileLeafRef,VendorID,Document_Type,Document_URL,Expiry_Date,Upload_Date,Status,Remarks&$orderby=Document_Type`;
        const libraryResponse: SPHttpClientResponse = await spHttpClient.get(
          libraryUrl,
          SPHttpClient.configurations.v1
        );

        if (libraryResponse.ok) {
          const libraryData = await libraryResponse.json();
          libraryDocs.push(...(libraryData.value || []).map((item: ISharePointListItem) => {
            const fileRef = getOptionalString(item.FileRef);
            const documentUrl = getHyperlinkUrl(item.Document_URL) || (fileRef ? `${pageContext.web.absoluteUrl}${fileRef}` : undefined);

            return {
              ID: item.ID,
              Title: getString(item.Title || item.FileLeafRef),
              VendorID: getString(item.VendorID),
              Document_Type: getString(item.Document_Type || item.Title || item.FileLeafRef),
              Document_URL: documentUrl,
              Expiry_Date: getOptionalString(item.Expiry_Date),
              Upload_Date: getOptionalString(item.Upload_Date),
              Status: getString(item.Status) || 'Under Review',
              Remarks: getString(item.Remarks),
            };
          }));
        }
      } catch (libraryError) {
        console.error('Error fetching vendor document library:', libraryError);
      }

      setDocuments([...listDocs, ...libraryDocs]);
    } catch (err) {
      console.error('Error fetching vendor documents:', err);
    } finally {
      setIsLoadingDocs(false);
    }
  }, [vendor, spHttpClient, pageContext]);

  React.useEffect(() => {
    if (isOpen && vendor) {
      fetchDocuments().catch(() => undefined);
      setUpdateMessage(undefined);
    }
  }, [isOpen, vendor, fetchDocuments]);

  // ─── Update Vendor Status ──────────────────────────────────────────────────

  const updateVendorStatus = async (newStatus: string): Promise<void> => {
    if (!vendor) return;
    setIsUpdating(true);
    setUpdateMessage(undefined);

    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items(${vendor.ID})`;
      const body: Record<string, unknown> = {
        Vendor_Status: newStatus,
      };

      if (newStatus === 'Active') {
        body.Approval_Date = new Date().toISOString();
      }

      const response: SPHttpClientResponse = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify(body),
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to update vendor: ${response.status}`);
      }

      setUpdateMessage({ type: MessageBarType.success, text: `Vendor status updated to "${newStatus}"` });
      setTimeout(() => {
        onVendorUpdated();
      }, 1000);
    } catch (err) {
      console.error('Error updating vendor:', err);
      setUpdateMessage({
        type: MessageBarType.error,
        text: err instanceof Error ? err.message : 'Failed to update vendor status',
      });
    } finally {
      setIsUpdating(false);
    }
  };

  // ─── Build Document Checklist ──────────────────────────────────────────────

  const getDocumentStatus = (docType: string): { status: string; document: IVendorDocument | undefined } => {
    const found = documents.find((d) => d.Document_Type === docType);
    if (!found) {
        return { status: 'Missing', document: undefined };
    }
    // Check if expired
    if (found.Expiry_Date && new Date(found.Expiry_Date) < new Date()) {
      return { status: 'Expired', document: found };
    }
    return { status: found.Status, document: found };
  };

  // ─── Render ────────────────────────────────────────────────────────────────

  if (!vendor) {
    return <></>;
  }

  const statusColors: Record<string, { bg: string; text: string }> = {
    Active: { bg: '#D1FAE5', text: '#065F46' },
    Inactive: { bg: '#F3F4F6', text: '#6B7280' },
    Pending: { bg: '#FEF3C7', text: '#92400E' },
    'Under Review': { bg: '#DBEAFE', text: '#1E40AF' },
    Approved: { bg: '#D1FAE5', text: '#065F46' },
    Rejected: { bg: '#FEE2E2', text: '#991B1B' },
    Suspended: { bg: '#FFEDD5', text: '#9A3412' },
    Blacklisted: { bg: '#FCE7F3', text: '#831843' },
  };

  const currentStatusColor = statusColors[vendor.Vendor_Status] || { bg: '#F3F4F6', text: '#374151' };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Vendor Details"
      closeButtonAriaLabel="Close"
    >
      <div className={classNames.panelContent}>
        {/* Update Message */}
        {updateMessage && (
          <MessageBar
            messageBarType={updateMessage.type}
            isMultiline={false}
            onDismiss={() => setUpdateMessage(undefined)}
            style={{ marginBottom: '12px' }}
          >
            {updateMessage.text}
          </MessageBar>
        )}

        {/* Vendor Header */}
        <div className={classNames.vendorHeader}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
            <div>
              <div className={classNames.vendorName}>{vendor.Vendor_Name}</div>
              <div className={classNames.vendorId}>{vendor.Vendor_Code}</div>
            </div>
            <span
              className={classNames.statusPill}
              style={{ backgroundColor: currentStatusColor.bg, color: currentStatusColor.text }}
            >
              {vendor.Vendor_Status}
            </span>
          </div>
        </div>

        {/* Vendor Info Grid */}
        <div className={classNames.infoGrid}>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Contact Person</span>
            <span className={classNames.infoValue}>{vendor.Contact_Person || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Email</span>
            <span className={classNames.infoValue}>{vendor.Email || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Phone</span>
            <span className={classNames.infoValue}>{vendor.Phone || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Category</span>
            <span className={classNames.infoValue}>{vendor.Vendor_Category || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Registration Type</span>
            <span className={classNames.infoValue}>{vendor.Registration_Type || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Tax ID (TIN)</span>
            <span className={classNames.infoValue}>{vendor.Tax_ID || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>RC Number (CAC)</span>
            <span className={classNames.infoValue}>{vendor.RC_Number || '—'}</span>
          </div>
          <div className={classNames.infoItem}>
            <span className={classNames.infoLabel}>Bank</span>
            <span className={classNames.infoValue}>{vendor.Bank_Name || '—'}</span>
          </div>
        </div>

        {vendor.Address && (
          <div style={{ marginBottom: '16px' }}>
            <span className={classNames.infoLabel}>Address</span>
            <div className={classNames.infoValue} style={{ marginTop: '2px' }}>{vendor.Address}</div>
          </div>
        )}

        <Separator />

        {/* Documents Checklist */}
        <div className={classNames.sectionTitle}>
          <Icon iconName="Documentation" style={{ marginRight: '6px' }} />
          Document Compliance ({documents.length}/{REQUIRED_DOCUMENTS.length - 1} uploaded)
        </div>

        {isLoadingDocs ? (
          <Spinner size={SpinnerSize.small} label="Loading documents..." />
        ) : (
          REQUIRED_DOCUMENTS.filter((d) => d !== 'Other').map((docType) => {
            const { status } = getDocumentStatus(docType);
            const statusStyle = DOC_STATUS_COLORS[status] || DOC_STATUS_COLORS.Missing;
            return (
              <div key={docType} className={classNames.docRow}>
                <span className={classNames.docName}>{docType}</span>
                <span
                  className={classNames.docStatus}
                  style={{ backgroundColor: statusStyle.bg, color: statusStyle.text }}
                >
                  <Icon iconName={statusStyle.icon} style={{ fontSize: 10 }} />
                  {status}
                </span>
              </div>
            );
          })
        )}

        {documents.length > 0 && (
          <>
            <div className={classNames.sectionTitle}>Uploaded Compliance Documents</div>
            {documents.map((document) => {
              const statusStyle = DOC_STATUS_COLORS[document.Status] || DOC_STATUS_COLORS.Missing;

              return (
                <div key={`${document.ID}-${document.Document_Type}`} className={classNames.docRow}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1 }}>
                    <span className={classNames.docName}>{document.Document_Type || document.Title}</span>
                    {document.Remarks && <Text variant="small" style={{ color: '#5A6A85' }}>{document.Remarks}</Text>}
                  </div>
                  <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                    <span
                      className={classNames.docStatus}
                      style={{ backgroundColor: statusStyle.bg, color: statusStyle.text }}
                    >
                      <Icon iconName={statusStyle.icon} style={{ fontSize: 10 }} />
                      {document.Status}
                    </span>
                    <DefaultButton
                      text="Preview"
                      iconProps={{ iconName: 'View' }}
                      disabled={!document.Document_URL}
                      onClick={() => {
                        if (document.Document_URL) {
                          window.open(document.Document_URL, '_blank');
                        }
                      }}
                      styles={{ root: { fontSize: 11, height: 28, padding: '0 8px' } }}
                    />
                  </Stack>
                </div>
              );
            })}
          </>
        )}

        <Separator />

        {/* Action Buttons */}
        <div className={classNames.sectionTitle}>Actions</div>
        <div className={classNames.actionButtons}>
          {vendor.Vendor_Status !== 'Active' && (
            <PrimaryButton
              text="Approve"
              iconProps={{ iconName: 'CheckMark' }}
              onClick={() => updateVendorStatus('Active')}
              disabled={isUpdating}
              styles={{ root: { backgroundColor: '#10B981', borderColor: '#10B981' } }}
            />
          )}
          {vendor.Vendor_Status !== 'Rejected' && vendor.Vendor_Status !== 'Blacklisted' && (
            <DefaultButton
              text="Reject"
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => updateVendorStatus('Rejected')}
              disabled={isUpdating}
              styles={{ root: { color: '#EF4444', borderColor: '#EF4444' } }}
            />
          )}
          {vendor.Vendor_Status === 'Active' && (
            <DefaultButton
              text="Suspend"
              iconProps={{ iconName: 'BlockContact' }}
              onClick={() => updateVendorStatus('Suspended')}
              disabled={isUpdating}
              styles={{ root: { color: '#F59E0B', borderColor: '#F59E0B' } }}
            />
          )}
        </div>

        {isUpdating && (
          <Stack horizontalAlign="center" style={{ marginTop: '12px' }}>
            <Spinner size={SpinnerSize.small} label="Updating..." />
          </Stack>
        )}

        {/* Notes */}
        {vendor.Notes && (
          <>
            <Separator />
            <div className={classNames.sectionTitle}>Notes</div>
            <Text variant="small" style={{ color: '#5A6A85' }}>
              {vendor.Notes}
            </Text>
          </>
        )}
      </div>
    </Panel>
  );
};

export default VendorDetailPanel;
