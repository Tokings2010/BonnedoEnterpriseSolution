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
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  TextField,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IProcurementRequest } from './BidManagementTab';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IBidDetailPanelProps {
  isOpen: boolean;
  request: IProcurementRequest | undefined;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  onDismiss: () => void;
  onBidUpdated: () => void;
}

interface IBidInvitation {
    ID: number;
    BidID: string;
    RequestID: string;
    VendorID: string;
    Vendor_Code: string;
    Vendor_Name: string;
    Invitation_Date: string;
    Due_Date: string;
    Status: string;
    Quote_Amount: number | undefined;
    Quote_Date: string | undefined;
    Quote_Validity_Days: number | undefined;
    Delivery_Days: number | undefined;
    Payment_Terms: string;
    Technical_Score: number | undefined;
    Commercial_Score: number | undefined;
    Notes: string;
}

interface IApprovedVendor {
  ID: number;
  Vendor_Code: string;
  Vendor_Name: string;
  Vendor_Category: string;
}

interface ISharePointListItem {
    ID: number;
    [key: string]: unknown;
}

const getString = (value: unknown): string => typeof value === 'string' ? value : '';
const getOptionalString = (value: unknown): string | undefined => typeof value === 'string' && value ? value : undefined;
const getNumber = (value: unknown): number | undefined => typeof value === 'number' ? value : undefined;

interface IQuoteForm {
    vendorBidId: number | undefined;
    quoteAmount: string;
    deliveryDays: string;
    paymentTerms: string;
    technicalScore: string;
    commercialScore: string;
}

// ─── Bid Status Colors ───────────────────────────────────────────────────────

const BID_STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  Invited: { bg: '#DBEAFE', text: '#1E40AF' },
  'Quote Submitted': { bg: '#EDE9FE', text: '#5B21B6' },
  Declined: { bg: '#FEE2E2', text: '#991B1B' },
  'No Response': { bg: '#F3F4F6', text: '#6B7280' },
  Awarded: { bg: '#D1FAE5', text: '#065F46' },
  'Not Awarded': { bg: '#F3F4F6', text: '#6B7280' },
};

// ─── Styles ──────────────────────────────────────────────────────────────────

const getStyles = (): ReturnType<typeof mergeStyleSets> =>
  mergeStyleSets({
    panelContent: {
      padding: '0 24px 24px',
    },
    requestHeader: {
      backgroundColor: '#F5F6FA',
      borderRadius: '8px',
      padding: '16px',
      marginBottom: '16px',
    },
    requestTitle: {
      fontSize: '16px',
      fontWeight: '700',
      color: '#1E2532',
      marginBottom: '4px',
    },
    requestId: {
      fontSize: '13px',
      fontFamily: 'Cascadia Code, Consolas, monospace',
      color: '#5A6A85',
    },
    infoRow: {
      display: 'flex',
      gap: '24px',
      flexWrap: 'wrap',
      marginTop: '12px',
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
    statusPill: {
      display: 'inline-block',
      padding: '2px 10px',
      borderRadius: '12px',
      fontSize: '11px',
      fontWeight: '600',
    },
    sourcingBadge: {
      display: 'inline-flex',
      alignItems: 'center',
      gap: '4px',
      padding: '4px 12px',
      borderRadius: '14px',
      fontSize: '12px',
      fontWeight: '600',
    },
    comparisonCard: {
      backgroundColor: '#F0FDF4',
      border: '1px solid #BBF7D0',
      borderRadius: '8px',
      padding: '12px',
      marginTop: '12px',
    },
    quoteFormContainer: {
      backgroundColor: '#FAFBFC',
      border: '1px solid #E5E7EB',
      borderRadius: '8px',
      padding: '16px',
      marginTop: '12px',
    },
    inviteSection: {
      marginTop: '12px',
      marginBottom: '12px',
    },
  });

// ─── Component ───────────────────────────────────────────────────────────────

const BidDetailPanel: React.FC<IBidDetailPanelProps> = ({
  isOpen,
  request,
  spHttpClient,
  pageContext,
  onDismiss,
  onBidUpdated,
}) => {
  const [bids, setBids] = React.useState<IBidInvitation[]>([]);
  const [approvedVendors, setApprovedVendors] = React.useState<IApprovedVendor[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
    const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [showQuoteForm, setShowQuoteForm] = React.useState(false);
  const [quoteForm, setQuoteForm] = React.useState<IQuoteForm>({
    vendorBidId: undefined,
    quoteAmount: '',
    deliveryDays: '',
    paymentTerms: '',
    technicalScore: '',
    commercialScore: '',
  });
  const [showInviteDropdown, setShowInviteDropdown] = React.useState(false);
  const [selectedVendorToInvite, setSelectedVendorToInvite] = React.useState<string>('');

  const classNames = getStyles() as unknown as { [key: string]: string };

  // ─── Fetch Bids ────────────────────────────────────────────────────────────

  const fetchBids = React.useCallback(async (): Promise<void> => {
    if (!request) return;
    setIsLoading(true);
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items?$filter=RequestID eq '${request.RequestID}'&$orderby=Quote_Amount asc`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) throw new Error(`Failed to fetch bids: ${response.status}`);

      const data = await response.json();
        const invitations: IBidInvitation[] = (data.value || []).map((item: ISharePointListItem) => ({
            ID: item.ID,
            BidID: getString(item.BidID),
            RequestID: getString(item.RequestID),
            VendorID: getString(item.VendorID || item.Vendor_Code),
            Vendor_Code: getString(item.Vendor_Code || item.VendorID),
            Vendor_Name: getString(item.Vendor_Name),
            Invitation_Date: getString(item.Invitation_Date),
            Due_Date: getString(item.Due_Date),
            Status: getString(item.Status) || 'Invited',
            Quote_Amount: getNumber(item.Quote_Amount),
            Quote_Date: getOptionalString(item.Quote_Date),
            Quote_Validity_Days: getNumber(item.Quote_Validity_Days),
            Delivery_Days: getNumber(item.Delivery_Days),
            Payment_Terms: getString(item.Payment_Terms),
            Technical_Score: getNumber(item.Technical_Score),
            Commercial_Score: getNumber(item.Commercial_Score),
            Notes: getString(item.Notes),
        }));

      setBids(invitations);
    } catch (err) {
      console.error('Error fetching bids:', err);
    } finally {
      setIsLoading(false);
    }
  }, [request, spHttpClient, pageContext]);

  // ─── Fetch Approved Vendors ────────────────────────────────────────────────

  const fetchApprovedVendors = React.useCallback(async (): Promise<void> => {
    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Vendors_Master')/items?$filter=Vendor_Status eq 'Active'&$select=ID,Vendor_Code,Vendor_Name,Vendor_Category&$top=200`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return;

      const data = await response.json();
        setApprovedVendors(
            (data.value || []).map((item: ISharePointListItem) => ({
                ID: item.ID,
                Vendor_Code: getString(item.Vendor_Code),
                Vendor_Name: getString(item.Vendor_Name),
                Vendor_Category: getString(item.Vendor_Category),
            }))
        );
    } catch (err) {
      console.error('Error fetching approved vendors:', err);
    }
  }, [spHttpClient, pageContext]);

  React.useEffect(() => {
    if (isOpen && request) {
        fetchBids().catch(() => undefined);
        fetchApprovedVendors().catch(() => undefined);
        setMessage(undefined);
      setShowQuoteForm(false);
      setShowInviteDropdown(false);
    }
  }, [isOpen, request, fetchBids, fetchApprovedVendors]);

  // ─── Invite Vendor ─────────────────────────────────────────────────────────

  const inviteVendor = async (): Promise<void> => {
    if (!request || !selectedVendorToInvite) return;
    setIsSubmitting(true);
    setMessage(undefined);

    const vendor = approvedVendors.find((v) => v.Vendor_Code === selectedVendorToInvite);
    if (!vendor) return;

    try {
      // Generate BidID (BID-YYYY-NNNN)
      const year = new Date().getFullYear().toString();
      const bidPrefix = `BID-${year}-`;
      const bidIdQueryUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items?$filter=startswith(BidID,'${bidPrefix}')&$orderby=BidID desc&$top=1&$select=BidID`;
      let nextBidNumber = 1;
      try {
        const bidIdResponse: SPHttpClientResponse = await spHttpClient.get(
          bidIdQueryUrl,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (bidIdResponse.ok) {
          const bidIdData = await bidIdResponse.json();
          if (bidIdData.value && bidIdData.value.length > 0) {
            const lastBidId = bidIdData.value[0].BidID || '';
            const lastBidNum = parseInt(lastBidId.split('-').pop() || '0', 10);
            if (!isNaN(lastBidNum)) nextBidNumber = lastBidNum + 1;
          }
        }
      } catch { /* first bid — start at 1 */ }

      const generatedBidID = `${bidPrefix}${nextBidNumber.toString().padStart(4, '0')}`;

      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items`;
        const body: Record<string, unknown> = {
            Title: `Bid for ${request.RequestID}`,
            BidID: generatedBidID,
            RequestID: request.RequestID,
            VendorID: vendor.Vendor_Code,
            Vendor_Name: vendor.Vendor_Name,
            Invitation_Date: new Date().toISOString(),
            Due_Date: request.Required_Date || new Date(Date.now() + 14 * 86400000).toISOString(),
            Status: 'Invited',
        };

      const response: SPHttpClientResponse = await spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
          body: JSON.stringify(body),
        }
      );

      if (!response.ok) throw new Error(`Failed to invite vendor: ${response.status}`);

        setMessage({ type: MessageBarType.success, text: `${vendor.Vendor_Name} invited successfully` });
      setSelectedVendorToInvite('');
      setShowInviteDropdown(false);
      await fetchBids();
    } catch (err) {
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to invite vendor' });
    } finally {
      setIsSubmitting(false);
    }
  };

  // ─── Submit Quote ──────────────────────────────────────────────────────────

  const submitQuote = async (): Promise<void> => {
    if (!quoteForm.vendorBidId || !quoteForm.quoteAmount) return;
    setIsSubmitting(true);
    setMessage(undefined);

    try {
      const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items(${quoteForm.vendorBidId})`;
        const body: Record<string, unknown> = {
            Status: 'Quote Submitted',
            Quote_Amount: parseFloat(quoteForm.quoteAmount),
            Quote_Date: new Date().toISOString(),
        };

      if (quoteForm.deliveryDays) body.Delivery_Days = parseInt(quoteForm.deliveryDays, 10);
      if (quoteForm.paymentTerms) body.Payment_Terms = quoteForm.paymentTerms;
      if (quoteForm.technicalScore) body.Technical_Score = parseInt(quoteForm.technicalScore, 10);
      if (quoteForm.commercialScore) body.Commercial_Score = parseInt(quoteForm.commercialScore, 10);

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

      if (!response.ok) throw new Error(`Failed to submit quote: ${response.status}`);

      setMessage({ type: MessageBarType.success, text: 'Quote recorded successfully' });
      setShowQuoteForm(false);
      setQuoteForm({ vendorBidId: undefined, quoteAmount: '', deliveryDays: '', paymentTerms: '', technicalScore: '', commercialScore: '' });
      await fetchBids();
    } catch (err) {
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to submit quote' });
    } finally {
      setIsSubmitting(false);
    }
  };

  // ─── Award Bid ─────────────────────────────────────────────────────────────

  const awardBid = async (bidId: number, vendorName: string): Promise<void> => {
    setIsSubmitting(true);
    setMessage(undefined);

    try {
      // 1. Set winning bid to "Awarded"
      const awardUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items(${bidId})`;
      await spHttpClient.post(awardUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
        },
        body: JSON.stringify({ Status: 'Awarded' }),
      });

      // 2. Set all other bids to "Not Awarded"
      for (const bid of bids) {
        if (bid.ID !== bidId && bid.Status !== 'Declined' && bid.Status !== 'No Response') {
          const url = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Bid_Invitations')/items(${bid.ID})`;
          await spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE',
            },
            body: JSON.stringify({ Status: 'Not Awarded' }),
          });
        }
      }

      // 3. Update procurement request status to "Awarded"
      if (request) {
        const reqUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Procurement_Requests')/items(${request.ID})`;
        await spHttpClient.post(reqUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({ Status: 'Awarded', Assigned_Vendor: vendorName }),
        });
      }

      setMessage({ type: MessageBarType.success, text: `Bid awarded to ${vendorName}` });
      setTimeout(() => onBidUpdated(), 1500);
    } catch (err) {
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to award bid' });
    } finally {
      setIsSubmitting(false);
    }
  };

  // ─── Bid Columns ───────────────────────────────────────────────────────────

  const bidColumns: IColumn[] = [
    {
      key: 'vendorName',
      name: 'Vendor',
      fieldName: 'Vendor_Name',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IBidInvitation) => (
        <Text style={{ fontWeight: 600 }}>{item.Vendor_Name}</Text>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 110,
      maxWidth: 130,
      onRender: (item: IBidInvitation) => {
        const colors = BID_STATUS_COLORS[item.Status] || { bg: '#F3F4F6', text: '#374151' };
        return (
          <span className={classNames.statusPill} style={{ backgroundColor: colors.bg, color: colors.text }}>
            {item.Status}
          </span>
        );
      },
    },
    {
      key: 'quoteAmount',
      name: 'Quote (₦)',
      fieldName: 'Quote_Amount',
      minWidth: 100,
      maxWidth: 130,
      onRender: (item: IBidInvitation) => (
        <Text style={{ fontWeight: item.Quote_Amount !== undefined ? 600 : 400 }}>
          {item.Quote_Amount !== undefined ? `₦${item.Quote_Amount.toLocaleString()}` : '—'}
        </Text>
      ),
    },
    {
      key: 'deliveryDays',
      name: 'Delivery',
      fieldName: 'Delivery_Days',
      minWidth: 70,
      maxWidth: 90,
      onRender: (item: IBidInvitation) => (
        <Text variant="small">{item.Delivery_Days !== undefined ? `${item.Delivery_Days} days` : '—'}</Text>
      ),
    },
    {
      key: 'technicalScore',
      name: 'Tech Score',
      fieldName: 'Technical_Score',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: IBidInvitation) => (
        <Text variant="small">{item.Technical_Score !== undefined ? `${item.Technical_Score}/100` : '—'}</Text>
      ),
    },
    {
      key: 'commercialScore',
      name: 'Comm Score',
      fieldName: 'Commercial_Score',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: IBidInvitation) => (
        <Text variant="small">{item.Commercial_Score !== undefined ? `${item.Commercial_Score}/100` : '—'}</Text>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 160,
      maxWidth: 200,
      onRender: (item: IBidInvitation) => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          {item.Status === 'Invited' && (
            <DefaultButton
              text="Enter Quote"
              iconProps={{ iconName: 'Edit' }}
              styles={{ root: { fontSize: '11px', height: '28px', padding: '0 8px' } }}
              onClick={() => {
                setQuoteForm({ ...quoteForm, vendorBidId: item.ID });
                setShowQuoteForm(true);
              }}
            />
          )}
          {item.Status === 'Quote Submitted' && request?.Status !== 'Awarded' && (
            <PrimaryButton
              text="Award"
              iconProps={{ iconName: 'Trophy2Solid' }}
              styles={{ root: { fontSize: '11px', height: '28px', padding: '0 8px', backgroundColor: '#10B981', borderColor: '#10B981' } }}
              onClick={() => awardBid(item.ID, item.Vendor_Name)}
            />
          )}
        </Stack>
      ),
    },
  ];

  // ─── Render ────────────────────────────────────────────────────────────────

    if (!request) {
        return <></>;
    }

  const sourcingColor = request.Sourcing_Method === 'Single Source'
    ? { bg: '#EDE9FE', text: '#7C3AED', icon: 'Contact' }
    : { bg: '#DBEAFE', text: '#2563EB', icon: 'People' };

  // Find best values for comparison
    const quotedBids = bids.filter((b): b is IBidInvitation & { Quote_Amount: number } => b.Quote_Amount !== undefined);
    const lowestQuote = quotedBids.length > 0 ? Math.min(...quotedBids.map((b) => b.Quote_Amount)) : undefined;
    const deliveryBids = quotedBids.filter((b): b is IBidInvitation & { Quote_Amount: number; Delivery_Days: number } => b.Delivery_Days !== undefined);
    const fastestDelivery = deliveryBids.length > 0
        ? Math.min(...deliveryBids.map((b) => b.Delivery_Days))
        : undefined;

  // Vendor dropdown options (exclude already invited)
  const invitedVendorIds = bids.map((b) => b.VendorID);
  const vendorOptions: IDropdownOption[] = approvedVendors
    .filter((v) => !invitedVendorIds.includes(v.Vendor_Code))
    .map((v) => ({ key: v.Vendor_Code, text: `${v.Vendor_Name} (${v.Vendor_Code})` }));

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.large}
      headerText="Procurement Request Details"
      closeButtonAriaLabel="Close"
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

        {/* Request Header */}
        <div className={classNames.requestHeader}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
            <div>
              <div className={classNames.requestTitle}>{request.Title}</div>
              <div className={classNames.requestId}>{request.RequestID}</div>
            </div>
            <span
              className={classNames.sourcingBadge}
              style={{ backgroundColor: sourcingColor.bg, color: sourcingColor.text }}
            >
              <Icon iconName={sourcingColor.icon} style={{ fontSize: 12 }} />
              {request.Sourcing_Method}
            </span>
          </div>

          <div className={classNames.infoRow}>
            <div className={classNames.infoItem}>
              <span className={classNames.infoLabel}>Project</span>
              <span className={classNames.infoValue}>{request.Project_Code || '—'}</span>
            </div>
            <div className={classNames.infoItem}>
              <span className={classNames.infoLabel}>Category</span>
              <span className={classNames.infoValue}>{request.Category || '—'}</span>
            </div>
            <div className={classNames.infoItem}>
              <span className={classNames.infoLabel}>Budget</span>
              <span className={classNames.infoValue}>
                {request.Budget_Estimate ? `₦${request.Budget_Estimate.toLocaleString()}` : '—'}
              </span>
            </div>
            <div className={classNames.infoItem}>
              <span className={classNames.infoLabel}>Required By</span>
              <span className={classNames.infoValue}>
                {request.Required_Date ? new Date(request.Required_Date).toLocaleDateString() : '—'}
              </span>
            </div>
            <div className={classNames.infoItem}>
              <span className={classNames.infoLabel}>Status</span>
              <span className={classNames.infoValue}>{request.Status}</span>
            </div>
          </div>
        </div>

        {/* Justification (for Single Source) */}
        {request.Sourcing_Method === 'Single Source' && request.Justification && (
          <div style={{ marginBottom: '16px' }}>
            <span className={classNames.infoLabel}>Single Source Justification</span>
            <Text variant="small" style={{ display: 'block', marginTop: '4px', color: '#374151' }}>
              {request.Justification}
            </Text>
          </div>
        )}

        <Separator />

        {/* Vendor Invitations / Bids */}
        <div className={classNames.sectionTitle}>
          <Icon iconName="People" style={{ marginRight: '6px' }} />
          Vendor Bids ({bids.length})
        </div>

        {isLoading ? (
          <Spinner size={SpinnerSize.small} label="Loading bids..." />
        ) : bids.length === 0 ? (
          <Text variant="small" style={{ color: '#5A6A85' }}>
            No vendors invited yet. Click &quot;Invite Vendor&quot; to start the bidding process.
          </Text>
        ) : (
          <>
            <DetailsList
              items={bids}
              columns={bidColumns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
              compact={true}
            />

            {/* Comparison Summary */}
            {quotedBids.length >= 2 && (
              <div className={classNames.comparisonCard}>
                <Text style={{ fontWeight: 600, fontSize: '13px', color: '#065F46' }}>
                  <Icon iconName="CompareUneven" style={{ marginRight: '4px' }} />
                  Comparison Summary
                </Text>
                <div style={{ display: 'flex', gap: '24px', marginTop: '8px' }}>
                    {lowestQuote !== undefined && (
                    <div>
                      <span className={classNames.infoLabel}>Lowest Quote</span>
                      <div style={{ fontWeight: 700, color: '#065F46' }}>₦{lowestQuote.toLocaleString()}</div>
                    </div>
                  )}
                    {fastestDelivery !== undefined && (
                    <div>
                      <span className={classNames.infoLabel}>Fastest Delivery</span>
                      <div style={{ fontWeight: 700, color: '#065F46' }}>{fastestDelivery} days</div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </>
        )}

        {/* Invite Vendor Section */}
        {request.Status !== 'Awarded' && request.Status !== 'Cancelled' && (
          <div className={classNames.inviteSection}>
            {!showInviteDropdown ? (
              <DefaultButton
                text="Invite Vendor"
                iconProps={{ iconName: 'AddFriend' }}
                onClick={() => setShowInviteDropdown(true)}
              />
            ) : (
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  placeholder="Select approved vendor..."
                  options={vendorOptions}
                  selectedKey={selectedVendorToInvite}
                  onChange={(_, option) => setSelectedVendorToInvite(option?.key as string || '')}
                  styles={{ root: { width: '300px' } }}
                />
                <PrimaryButton
                  text="Send Invite"
                  onClick={inviteVendor}
                  disabled={!selectedVendorToInvite || isSubmitting}
                />
                <DefaultButton text="Cancel" onClick={() => setShowInviteDropdown(false)} />
              </Stack>
            )}
          </div>
        )}

        {/* Quote Entry Form */}
        {showQuoteForm && (
          <div className={classNames.quoteFormContainer}>
            <Text style={{ fontWeight: 600, marginBottom: '12px', display: 'block' }}>
              Enter Vendor Quote
            </Text>
            <Stack tokens={{ childrenGap: 10 }}>
              <TextField
                label="Quote Amount (₦)"
                value={quoteForm.quoteAmount}
                onChange={(_, v) => setQuoteForm({ ...quoteForm, quoteAmount: v || '' })}
                type="number"
                required
              />
              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <TextField
                  label="Delivery Days"
                  value={quoteForm.deliveryDays}
                  onChange={(_, v) => setQuoteForm({ ...quoteForm, deliveryDays: v || '' })}
                  type="number"
                  styles={{ root: { flex: 1 } }}
                />
                <TextField
                  label="Payment Terms"
                  value={quoteForm.paymentTerms}
                  onChange={(_, v) => setQuoteForm({ ...quoteForm, paymentTerms: v || '' })}
                  placeholder="e.g. Net 30"
                  styles={{ root: { flex: 1 } }}
                />
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <TextField
                  label="Technical Score (0-100)"
                  value={quoteForm.technicalScore}
                  onChange={(_, v) => setQuoteForm({ ...quoteForm, technicalScore: v || '' })}
                  type="number"
                  styles={{ root: { flex: 1 } }}
                />
                <TextField
                  label="Commercial Score (0-100)"
                  value={quoteForm.commercialScore}
                  onChange={(_, v) => setQuoteForm({ ...quoteForm, commercialScore: v || '' })}
                  type="number"
                  styles={{ root: { flex: 1 } }}
                />
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton text="Submit Quote" onClick={submitQuote} disabled={isSubmitting || !quoteForm.quoteAmount} />
                <DefaultButton text="Cancel" onClick={() => setShowQuoteForm(false)} />
              </Stack>
            </Stack>
          </div>
        )}

        {isSubmitting && (
          <Stack horizontalAlign="center" style={{ marginTop: '12px' }}>
            <Spinner size={SpinnerSize.small} label="Processing..." />
          </Stack>
        )}
      </div>
    </Panel>
  );
};

export default BidDetailPanel;
