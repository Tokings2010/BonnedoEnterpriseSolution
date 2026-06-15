import * as React from 'react';
import {
  Panel,
  PanelType,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  getTheme,
  mergeStyleSets,
  ScrollablePane,
  ScrollbarVisibility,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';
import { PurchaseOrderPdfService } from '../services/PurchaseOrderPdfService';
import { IApprovalRecord, IApprovalAction } from '../models/DataModels';

export interface IApprovalPanelProps {
  isOpen: boolean;
  record: IApprovalRecord | undefined;
  listName: string;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  onDismiss: () => void;
  onApprovalComplete?: (record: IApprovalRecord) => void;
}

interface IApprovalPanelState {
  comment: string;
  isLoading: boolean;
  isSubmitting: boolean;
  error: string | undefined;
  successMessage: string | undefined;
  approvalHistory: IApprovalAction[];
}

const ApprovalPanel: React.FC<IApprovalPanelProps> = ({
  isOpen,
  record,
  listName,
  spHttpClient,
  pageContext,
  userDisplayName,
  onDismiss,
  onApprovalComplete,
}) => {
  const [state, setState] = React.useState<IApprovalPanelState>({
    comment: '',
    isLoading: false,
    isSubmitting: false,
    error: undefined,
    successMessage: undefined,
    approvalHistory: [],
  });

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  // Parse approval history when record changes
  React.useEffect(() => {
    if (record && record.Approval_History) {
      try {
        const history = JSON.parse(record.Approval_History);
        setState((prev) => ({
          ...prev,
          approvalHistory: Array.isArray(history) ? history : [],
          comment: '',
          error: undefined,
          successMessage: undefined,
        }));
      } catch (error) {
        console.error('Error parsing approval history:', error);
        setState((prev) => ({
          ...prev,
          approvalHistory: [],
          comment: '',
        }));
      }
    }
  }, [record, isOpen]);

  // Handle approve action
  const handleApprove = React.useCallback(async () => {
    if (!record) return;

    setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

    try {
      const timestamp = new Date().toISOString();
      const approvalAction: IApprovalAction = {
        action: 'approve',
        comment: state.comment,
        timestamp,
        approver: userDisplayName,
      };

      const updatedHistory = [...state.approvalHistory, approvalAction];

      // Material Request: Use multi-stage approval logic
      if (listName === 'PRC_Material_Request_Register') {
        await sharePointService.processMaterialRequestApprovalAction(
          record.ID,
          'approve',
          userDisplayName || '',
          state.comment
        );
      } else {
        // Update SharePoint list (default behavior for other modules)
        await sharePointService.updateListItem(listName, record.ID, {
          Approval_Status: 'Approved',
          Current_Approver: userDisplayName,
          Approval_History: JSON.stringify(updatedHistory),
        });
      }

      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        successMessage: listName === 'PRC_Material_Request_Register' 
          ? 'Approval recorded. Checking next stage...' 
          : 'Invoice approved successfully!',
        comment: '',
      }));

      // Notify parent component
      if (onApprovalComplete) {
        const updatedRecord: IApprovalRecord = {
          ...record,
          Approval_Status: 'Approved',
          Current_Approver: userDisplayName,
          Approval_History: JSON.stringify(updatedHistory),
        };
        onApprovalComplete(updatedRecord);
      }

      // Auto-generate PDF for Purchase Orders on final approval
      if (listName === 'PRC_Purchase_Order_Register') {
        try {
          const pdfService = new PurchaseOrderPdfService(spHttpClient, pageContext);
          await pdfService.generateAndDistributePO({
            PO_Number: record.Title,
            Vendor: record.Vendor,
            TotalAmount: record.TotalAmount || record.Amount || 0,
          });
        } catch (pdfError) {
          console.error('PDF generation failed after PO approval:', pdfError);
        }
      }

      // Finalize Payment Request approval + notify vendor
      if (listName === 'FIN_Payment_Request_Register') {
        try {
          await sharePointService.finalizePaymentRequestApproval(record.ID);
          console.log(`[Payment Request] Finalized approval for: ${record.Title}`);
          // TODO: Send email confirmation to vendor using NotificationService + MSGraphClient
        } catch (prError) {
          console.error('Failed to finalize Payment Request approval:', prError);
        }
      }

      // Close panel after success
      setTimeout(() => {
        onDismiss();
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to approve invoice';
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [record, state.comment, state.approvalHistory, userDisplayName, sharePointService, listName, onApprovalComplete, onDismiss]);

  // Handle reject action
  const handleReject = React.useCallback(async () => {
    if (!record) return;

    // Validate comment is provided for rejection
    if (!state.comment.trim()) {
      setState((prev) => ({
        ...prev,
        error: 'Comment is required when rejecting an invoice.',
      }));
      return;
    }

    setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

    try {
      const timestamp = new Date().toISOString();
      const approvalAction: IApprovalAction = {
        action: 'reject',
        comment: state.comment,
        timestamp,
        approver: userDisplayName,
      };

      const updatedHistory = [...state.approvalHistory, approvalAction];

      // Update SharePoint list
      await sharePointService.updateListItem(listName, record.ID, {
        Approval_Status: 'Rejected',
        Current_Approver: userDisplayName,
        Approval_History: JSON.stringify(updatedHistory),
      });

      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        successMessage: 'Invoice rejected successfully!',
        comment: '',
      }));

      // Notify parent component
      if (onApprovalComplete) {
        const updatedRecord: IApprovalRecord = {
          ...record,
          Approval_Status: 'Rejected',
          Current_Approver: userDisplayName,
          Approval_History: JSON.stringify(updatedHistory),
        };
        onApprovalComplete(updatedRecord);
      }

      // Close panel after success
      setTimeout(() => {
        onDismiss();
      }, 1500);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to reject invoice';
      setState((prev) => ({
        ...prev,
        isSubmitting: false,
        error: errorMessage,
      }));
    }
  }, [record, state.comment, state.approvalHistory, userDisplayName, sharePointService, listName, onApprovalComplete, onDismiss]);

  const theme = getTheme();
  const classNames = mergeStyleSets({
    panelContent: {
      padding: '20px',
    },
    section: {
      marginBottom: '24px',
    },
    fieldLabel: {
      fontWeight: 600,
      marginBottom: '8px',
      color: theme.palette.neutralPrimary,
    },
    fieldValue: {
      color: theme.palette.neutralSecondary,
      marginBottom: '16px',
    },
    historyContainer: {
      backgroundColor: theme.palette.neutralLighterAlt,
      padding: '12px',
      borderRadius: '4px',
      maxHeight: '200px',
      overflow: 'auto',
    },
    historyItem: {
      padding: '8px',
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      fontSize: '12px',
      marginBottom: '8px',
      selectors: {
        '&:last-child': {
          borderBottom: 'none',
        },
      },
    },
    actionBadge: {
      display: 'inline-block',
      padding: '2px 8px',
      borderRadius: '3px',
      fontSize: '11px',
      fontWeight: 600,
      marginRight: '8px',
    },
    approveBadge: {
      backgroundColor: theme.palette.green,
      color: theme.palette.white,
    },
    rejectBadge: {
      backgroundColor: theme.palette.red,
      color: theme.palette.white,
    },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
    },
    commentBox: {
      marginTop: '12px',
    },
  });

  if (!record) {
    return null;
  }

  const isDisabled = state.isSubmitting;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Invoice Approval"
      closeButtonAriaLabel="Close"
      isLightDismiss={!state.isSubmitting}
    >
      <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
        <div className={classNames.panelContent}>
          {/* Error Message */}
          {state.error && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {state.error}
              </MessageBar>
            </div>
          )}

          {/* Success Message */}
          {state.successMessage && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.success} isMultiline>
                {state.successMessage}
              </MessageBar>
            </div>
          )}

          {/* Invoice Details Section */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
              Invoice Details
            </Text>

            {/* Invoice Number */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Invoice Number</Label>
              <Text className={classNames.fieldValue}>
                {record.InvoiceNumber || record.Title || 'N/A'}
              </Text>
            </div>

            {/* Vendor */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Vendor</Label>
              <Text className={classNames.fieldValue}>
                {record.Vendor || 'N/A'}
              </Text>
            </div>

            {/* Amount */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Amount</Label>
              <Text className={classNames.fieldValue}>
                {record.Amount ? `$${record.Amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : 'N/A'}
              </Text>
            </div>

            {/* Fund Type */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Fund Type</Label>
              <Text className={classNames.fieldValue}>
                {record.FundType || 'N/A'}
              </Text>
            </div>

            {/* Status */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Current Status</Label>
              <Text
                className={classNames.fieldValue}
                style={{
                  color:
                    record.Approval_Status === 'Approved'
                      ? theme.palette.green
                      : record.Approval_Status === 'Rejected'
                        ? theme.palette.red
                        : theme.palette.orange,
                  fontWeight: 600,
                }}
              >
                {record.Approval_Status || 'Pending'}
              </Text>
            </div>
          </div>

          <Separator />

          {/* Approval History Section */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              Approval History
            </Text>

            {state.approvalHistory.length > 0 ? (
              <div className={classNames.historyContainer}>
                {state.approvalHistory.map((action, index) => (
                  <div key={index} className={classNames.historyItem}>
                    <div style={{ marginBottom: '4px' }}>
                      <span
                        className={`${classNames.actionBadge} ${action.action === 'approve'
                          ? classNames.approveBadge
                          : classNames.rejectBadge
                          }`}
                      >
                        {action.action.toUpperCase()}
                      </span>
                      <span style={{ color: theme.palette.neutralSecondary }}>
                        by {action.approver}
                      </span>
                    </div>
                    <div style={{ color: theme.palette.neutralTertiary, fontSize: '11px', marginBottom: '4px' }}>
                      {new Date(action.timestamp).toLocaleString()}
                    </div>
                    {action.comment && (
                      <div style={{ color: theme.palette.neutralSecondary, fontStyle: 'italic' }}>
                        &quot;{action.comment}&quot;
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ) : (
              <Text className={classNames.fieldValue}>No approval history yet</Text>
            )}
          </div>

          <Separator />

          {/* Comment Section */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              Your Decision
            </Text>

            <TextField
              label="Comment"
              placeholder="Add a comment (required for rejection)"
              multiline
              rows={4}
              value={state.comment}
              onChange={(_, newValue) =>
                setState((prev) => ({ ...prev, comment: newValue || '' }))
              }
              disabled={isDisabled}
              className={classNames.commentBox}
            />

            {/* Action Buttons */}
            <div className={classNames.buttonGroup}>
              <PrimaryButton
                text="Approve"
                onClick={handleApprove}
                disabled={isDisabled}
                iconProps={{ iconName: 'CheckMark' }}
              />
              <DefaultButton
                text="Reject"
                onClick={handleReject}
                disabled={isDisabled}
                iconProps={{ iconName: 'Cancel' }}
              />
            </div>
          </div>

          {/* Loading Indicator */}
          {state.isSubmitting && (
            <div style={{ marginTop: '24px', textAlign: 'center' }}>
              <Spinner size={SpinnerSize.medium} label="Processing your decision..." />
            </div>
          )}
        </div>
      </ScrollablePane>
    </Panel>
  );
};

export default ApprovalPanel;
