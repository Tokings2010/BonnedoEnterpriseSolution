import * as React from 'react';
import {
  Panel,
  PanelType,
  Text,
  Image,
  PrimaryButton,
  DefaultButton,
  Separator,
  Label,
  TextField,
  getTheme,
  mergeStyleSets,
  ScrollablePane,
  ScrollbarVisibility,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SharePointService } from '../services/SharePointService';
import { NotificationService } from '../services/NotificationService';

export interface IProcurementDetailsPanelProps {
  isOpen: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  record: any | undefined;
  recordType: 'MR' | 'PR' | 'PO' | 'GRN';
  onDismiss: () => void;
  onApprove?: () => void;
  onReject?: () => void;
  onConvert?: (newRecordType: 'MR' | 'PR' | 'PO' | 'GRN') => void;
  showApprovalButtons?: boolean;
  spHttpClient?: SPHttpClient;
  pageContext?: PageContext;
  onRefresh?: () => void;
  currentUserDisplayName?: string;
  onTrackApproval?: () => void;
  webPartContext?: WebPartContext;
}

const ProcurementDetailsPanel: React.FC<IProcurementDetailsPanelProps> = ({
  isOpen,
  record,
  recordType,
  onDismiss,
  onApprove,
  onReject,
  onConvert,
  showApprovalButtons = true,
  spHttpClient,
  pageContext,
  onRefresh,
  currentUserDisplayName,
  onTrackApproval,
  webPartContext,
}) => {
  const theme = getTheme();

  // Helper to extract display value from lookup objects (expanded via $expand)
  const getDisplayValue = (value: any): string => {
    if (value === null || value === undefined) return 'N/A';
    if (typeof value === 'object') {
      return value.Title || value.Name || value.Project_Code || value.Material_Name || String(value.ID || '');
    }
    return String(value);
  };

  // Determine if current user can approve/reject
  const canApprove = React.useMemo(() => {
    if (!record || !currentUserDisplayName) return false;
    
    const currentApprover = record.Current_Approver || record.Approver || record.ApproverId;
    if (!currentApprover) return false;

    // Normalize approver value
    let approverStr = '';
    if (typeof currentApprover === 'object') {
      approverStr = currentApprover.Title || currentApprover.Email || currentApprover.LoginName || '';
    } else {
      approverStr = String(currentApprover);
    }

    const userLower = currentUserDisplayName.toLowerCase().trim();
    const approverLower = approverStr.toLowerCase().trim();

    // Match by display name, email prefix, or contains
    return approverLower === userLower ||
           approverLower.includes(userLower) ||
           userLower.includes(approverLower) ||
            approverLower.split('@')[0] === userLower.split('@')[0];
  }, [record, currentUserDisplayName]);

  // Reusable helper to safely extract numeric ID from lookup fields (Person, Lookup, etc.)
  const getLookupId = (val: any): number | undefined => {
    if (typeof val === 'number') return val;
    if (val && typeof val === 'object') return val.ID || val.Id || val.id;
    return undefined;
  };

  const [isProcessing, setIsProcessing] = React.useState(false);
  const [showRejectComment, setShowRejectComment] = React.useState(false);
  const [rejectComment, setRejectComment] = React.useState('');
  const [approveComment, setApproveComment] = React.useState('');
  const [error, setError] = React.useState<string | undefined>();
  const [successMessage, setSuccessMessage] = React.useState<string | undefined>();
  const [hasExistingConversion, setHasExistingConversion] = React.useState(false);

  // State for focused PO conversion dialog (prompt for Vendor + Delivery Date when converting PR → PO)
  // This follows ERP best practice: capture critical fields at conversion time for complete PO reporting
  const [isPOVendorDialogOpen, setIsPOVendorDialogOpen] = React.useState(false);
  const [poDialogVendors, setPoDialogVendors] = React.useState<any[]>([]);
  const [poDialogSelectedVendorId, setPoDialogSelectedVendorId] = React.useState<number | undefined>(undefined);
  const [poDialogDeliveryDate, setPoDialogDeliveryDate] = React.useState<string>('');
  const [isLoadingVendorsForDialog, setIsLoadingVendorsForDialog] = React.useState(false);

  // Create SharePointService instance
  const sharePointService = React.useMemo(() => {
    if (spHttpClient && pageContext) {
      return new SharePointService(spHttpClient, pageContext);
    }
    return null;
  }, [spHttpClient, pageContext]);

  // Check if MR has already been converted to PR by querying PR list
  React.useEffect(() => {
    const checkExistingConversion = async () => {
      if (!sharePointService || recordType !== 'MR' || !record?.ID) {
        setHasExistingConversion(false);
        return;
      }
      try {
        const prListName = 'PRC_Purchase_Requisition_Register';
        const filter = `Material_RequestId eq ${record.ID}`;
        const existing = await sharePointService.getListData(prListName, filter, 1);
        setHasExistingConversion(existing && existing.length > 0);
      } catch {
        setHasExistingConversion(false);
      }
    };
    checkExistingConversion();
  }, [record, recordType, sharePointService]);

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
      wordBreak: 'break-word',
    },
    qrCodeContainer: {
      textAlign: 'center',
      padding: '16px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '4px',
      marginBottom: '16px',
    },
    qrImage: {
      maxWidth: '150px',
      height: 'auto',
    },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
      flexWrap: 'wrap',
    },
    headerContainer: {
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      marginBottom: '4px',
    },
    statusBadge: {
      display: 'inline-flex',
      alignItems: 'center',
      padding: '4px 12px',
      borderRadius: '12px',
      fontSize: '12px',
      fontWeight: 600,
      textTransform: 'uppercase',
    },
    recordId: {
      fontSize: '14px',
      color: theme.palette.neutralSecondary,
    },
    convertButton: {
      backgroundColor: theme.palette.themeLight,
    },
    rejectSection: {
      marginTop: '16px',
      padding: '16px',
      backgroundColor: theme.palette.neutralLighterAlt,
      borderRadius: '4px',
    },
  });

  if (!record) {
    return null;
  }

  // Generate QR Code using external API
  const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(record.ID?.toString() || record.Title || '')}`;

  const getRecordTypeLabel = (): string => {
    switch (recordType) {
      case 'MR':
        return 'Material Request';
      case 'PR':
        return 'Purchase Requisition';
      case 'PO':
        return 'Purchase Order';
      case 'GRN':
        return 'Goods Received Note';
      default:
        return 'Record';
    }
  };

  const getStatusColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'approved':
        return theme.palette.green;
      case 'rejected':
        return theme.palette.red;
      case 'pending':
        return theme.palette.orange;
      case 'draft':
        return theme.palette.neutralSecondary;
      case 'active':
        return theme.palette.green;
      case 'completed':
        return theme.palette.blue;
      default:
        return theme.palette.neutralSecondary;
    }
  };

  const getStatusBackgroundColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'approved':
        return '#dff6dd';
      case 'rejected':
        return '#fde7e9';
      case 'pending':
        return '#fff4ce';
      case 'draft':
        return theme.palette.neutralLighter;
      case 'active':
        return '#dff6dd';
      case 'completed':
        return '#e6f2ff';
      default:
        return theme.palette.neutralLighter;
    }
  };

  const getStatusTextColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'approved':
        return '#107c10';
      case 'rejected':
        return '#a80000';
      case 'pending':
        return '#ca5010';
      case 'draft':
        return theme.palette.neutralDark;
      case 'active':
        return theme.palette.greenDark;
      case 'completed':
        return theme.palette.blueDark;
      default:
        return theme.palette.neutralDark;
    }
  };

  // Get the list name for the current record type
  const getListName = (): string => {
    switch (recordType) {
      case 'MR':
        return 'PRC_Material_Request_Register';
      case 'PR':
        return 'PRC_Purchase_Requisition_Register';
      case 'PO':
        return 'PRC_Purchase_Order_Register';
      case 'GRN':
        return 'PRC_GRN_Register';
      default:
        return '';
    }
  };

  // Handle Approve
  const handleApprove = async (): Promise<void> => {
    if (!sharePointService) {
      setError('SharePoint service not available');
      return;
    }

    setIsProcessing(true);
    setError(undefined);

    try {
      const approverEmail = currentUserDisplayName || 'Current User';
      const comment = approveComment || '';

      if (recordType === 'MR') {
        await sharePointService.processMaterialRequestApprovalAction(record.ID, 'approve', approverEmail, comment);
      } else if (recordType === 'PR') {
        await sharePointService.processPurchaseRequisitionApprovalAction(record.ID, 'approve', approverEmail, comment);
      } else if (recordType === 'GRN') {
        // GRN 1-stage approval + trigger Finance Manager notification (from Payment Request matrix)
        const result = await sharePointService.processGoodsReceivedNoteApprovalAction(record.ID, 'approve', approverEmail, comment);

        if (result?.financeManagerEmail && webPartContext) {
          try {
            const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
            const notificationService = new NotificationService(graphClient);

            const projectName = record.Project_Code?.Title || record.Project_Code?.Project_Code || 'Unknown Project';
            const deepLink = `${window.location.origin}${window.location.pathname}?grn=${record.ID}`;

            await notificationService.sendMaterialRequestApprovalNotification({
              requestTitle: `GRN ${record.Title || record.ID} - Finance Action Required`,
              project: projectName,
              amount: record.Total_Value || record.Quantity_Received || 0,
              approverEmail: result.financeManagerEmail,
              deepLink,
            });

            console.log('[GRN Approval] Finance Manager notification sent to:', result.financeManagerEmail);
          } catch (notifyErr) {
            console.error('[GRN] Failed to send Finance Manager email notification:', notifyErr);
          }
        }
      } else {
        // fallback simple approve for PO/others
        await sharePointService.updateListItem(getListName(), record.ID, {
          Approval_Status: 'Approved',
          Status: 'Approved',
          Approval_Completed_On: new Date().toISOString()
        });
      }

      // After each stage approval (for MR/PR/PO), notify the *next* approver with deep link (if any)
      // PO uses amount-threshold approval (unique) but still benefits from the same notify pattern if multi-stage
      if (webPartContext && (recordType === 'MR' || recordType === 'PR' || recordType === 'PO')) {
        try {
          const expand = 'Project_Code';
          const updatedItems = await sharePointService.getListData(getListName(), `ID eq ${record.ID}`, 1, 0, expand);
          if (updatedItems.length > 0) {
            const fresh = updatedItems[0];
            const nextApproverField = fresh.Current_Approver;
            const nextApprover = typeof nextApproverField === 'object'
              ? (nextApproverField.Email || nextApproverField.Title || nextApproverField.LoginName)
              : nextApproverField;
            if (nextApprover) {
              const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
              const notificationService = new NotificationService(graphClient);
              const projectName = fresh.Project_Code?.Project_Code || fresh.Project_Code?.Title || (typeof fresh.Project_Code === 'string' ? fresh.Project_Code : 'Unknown Project');
              const deepLink = `${window.location.origin}${window.location.pathname}?${recordType.toLowerCase()}=${record.ID}`;
              const reqTitle = fresh.Title || fresh.Request_No || fresh.Description || `${recordType}-${record.ID}`;

              console.log(`[${recordType}] Stage approval notification to next approver:`, nextApprover);

              await notificationService.sendMaterialRequestApprovalNotification({
                requestTitle: reqTitle,
                project: projectName,
                amount: fresh.Estimated_Cost ?? fresh.EstimatedCost,
                approverEmail: nextApprover,
                deepLink,
              });
            }
          }
        } catch (notifyErr) {
          console.error(`[${recordType}] Stage notification failed (approval still recorded):`, notifyErr);
        }
      }

      setSuccessMessage(`${getRecordTypeLabel()} approved successfully!`);

      if (onApprove) {
        onApprove();
      }

      if (onRefresh) {
        setTimeout(() => {
          onRefresh();
          onDismiss();
        }, 1500);
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to approve record';
      setError(errorMessage);
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle Reject with comment
  const handleReject = async (): Promise<void> => {
    if (!sharePointService) {
      setError('SharePoint service not available');
      return;
    }

    if (!rejectComment.trim()) {
      setError('Please provide a reason for rejection');
      return;
    }

    setIsProcessing(true);
    setError(undefined);

    try {
      const updates: { [key: string]: unknown } = {
        Approval_Status: 'Rejected',
        Status: 'Rejected',
        Rejection_Reason: rejectComment,
      };

      await sharePointService.updateListItem(getListName(), record.ID, updates);

      setSuccessMessage(`${getRecordTypeLabel()} rejected.`);

      if (onReject) {
        onReject();
      }

      if (onRefresh) {
        setTimeout(() => {
          onRefresh();
          onDismiss();
        }, 1500);
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to reject record';
      setError(errorMessage);
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle Convert (workflow progression)
  // poOverrides is used only for PO creation from PR when we prompt the user for missing Vendor/Delivery Date in a focused dialog (ERP best practice)
  const handleConvert = async (
    targetType: 'MR' | 'PR' | 'PO' | 'GRN',
    poOverrides?: { vendorId?: number; deliveryDate?: string | null }
  ): Promise<void> => {
    if (!sharePointService) {
      setError('SharePoint service not available');
      return;
    }

    setIsProcessing(true);
    setError(undefined);

    try {
      // Best-effort: mark source as completed to prevent duplicates.
      // We wrap this in try/catch because:
      // - The list may not have the Approval_Status/Status columns yet
      // - Concurrency (409 Conflict) can occur if the item was just modified
      // - For the new PR→PO flow we no longer call this here at all (the PO form does it after successful submit)
      try {
        await sharePointService.updateListItem(getListName(), record.ID, {
          Approval_Status: 'Approved',
          Status: 'Completed',
        });
      } catch (markErr) {
        console.warn('[Convert] Could not mark source record as Completed (non-fatal):', markErr);
        // Continue anyway – creating the target record is more important for the user
      }

      // Then create the new record
      let newListName = '';
      let newItemData: { [key: string]: unknown } = {};

      const generatePRCode = async (): Promise<string> => {
        try {
          const allItems = await sharePointService.getListData('PRC_Purchase_Requisition_Register', undefined, 500, 0, undefined);
          let nextNumber = 1;
          if (allItems && allItems.length > 0) {
            allItems.forEach((item) => {
              const title = item.Title || item.PR_Number || '';
              const matches = title.match(/\d+$/);
              if (matches) {
                const num = parseInt(matches[0], 10);
                if (num >= nextNumber) nextNumber = num + 1;
              }
            });
          }
          return `PR-${nextNumber.toString().padStart(4, '0')}`;
        } catch {
          return `PR-${Date.now().toString().slice(-6)}`;
        }
      };

      switch (targetType) {
        case 'PR':
          newListName = 'PRC_Purchase_Requisition_Register';
          const prCode = await generatePRCode();
          newItemData = {
            Title: prCode,
            PR_Number: prCode,
            Description: typeof record.Description === 'string' ? record.Description : (record.Description?.Title || record.Material?.Title || record.Title || ''),
            Quantity: typeof record.Quantity === 'number' ? record.Quantity : (parseInt(record.Quantity) || 1),
            Estimated_Cost: record.Estimated_Cost || 0,
            Status: 'Draft',
            Project_CodeId: getLookupId(record.Project_CodeId) || getLookupId(record.Project_Code),
            Material_RequestId: record.ID,
          };
          break;
        case 'PO':
          newListName = 'PRC_Purchase_Order_Register';
          const poCode = `BON-PO-${Date.now().toString().slice(-6)}`;
          // Robust amount extraction (handle various casings from grid/details records)
          const poAmount = Number(
            record.Estimated_Cost ?? record.EstimatedCost ?? record.TotalAmount ?? record.Amount ??
            record.totalAmount ?? record.estimatedCost ?? record['Estimated Cost'] ?? 0
          ) || 0;
          newItemData = {
            Title: poCode,
            PO_Number: poCode,
            PR_NumberId: record.ID,
            // Multi-PR support: we only write to PR_ReferencesId if the column exists as multi-lookup.
            // For now we rely on the legacy PR_NumberId (single) until the column is created.
            // PR_ReferencesId: { results: [record.ID] },
            Project_CodeId: getLookupId(record.Project_CodeId) || getLookupId(record.Project_Code),
            VendorId: poOverrides?.vendorId ?? (getLookupId(record.VendorId) || getLookupId(record.Vendor)),
            Amount: poAmount,
            TotalAmount: poAmount,
            Delivery_Date: poOverrides?.deliveryDate ?? (record.Delivery_Date || record.Required_Date || record.DeliveryDate || null),
            Status: 'Draft',
          };
          break;
        case 'GRN':
          // STEP 2: Validate PO is Approved before allowing GRN creation (same rule as the GRN form)
          const poStatusForGRN = record.Approval_Status || record.Status;
          if (poStatusForGRN && poStatusForGRN !== 'Approved') {
            setError('Cannot create GRN: The linked Purchase Order must be Approved first.');
            setIsProcessing(false);
            return;
          }

          newListName = 'PRC_GRN_Register';
          const grnCode = `GRN-${Date.now().toString().slice(-6)}`;

          // Pull Vendor from the PO if available (for complete GRN data)
          const grnVendorId = getLookupId(record.VendorId) || getLookupId(record.Vendor);

          newItemData = {
            Title: grnCode,
            GRN_Number: grnCode,
            PO_NumberId: record.ID,
            // Safe columns only — see GoodsReceivedNoteModule.tsx for full recommended list
            Quantity_Received: record.Quantity || 1,
            VendorId: grnVendorId,
            // Removed Related_PO_ID / Related_PO_Title — they do not exist on PRC_GRN_Register
          };
          break;
      }

      // Deep sanitize + correct field names for PRC_Purchase_Requisition_Register
      const sanitizeValue = (v: any): any => {
        if (v === null || v === undefined) return null;
        if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
        if (typeof v === 'object') {
          if (v.Title) return String(v.Title);
          if (v.Name) return String(v.Name);
          if (v.ID || v.Id) return v.ID || v.Id;
          return null;
        }
        return String(v);
      };

      Object.keys(newItemData).forEach(key => {
        newItemData[key] = sanitizeValue(newItemData[key]);
        if (newItemData[key] && typeof newItemData[key] === 'object') {
          delete newItemData[key];
        }
      });

      const createdItem = await sharePointService.createListItem(newListName, newItemData);

      if (targetType === 'PR' && createdItem && createdItem.ID) {
        await sharePointService.initializePurchaseRequisitionApproval(createdItem.ID);
      } else if (targetType === 'PO' && createdItem && createdItem.ID) {
        // Initialize the unique PO amount/threshold approval (Stage 1 or CFO Stage 2)
        const poAmountForApproval = Number((newItemData.Amount as any) ?? (newItemData.TotalAmount as any) ?? 0);
        await sharePointService.initializePurchaseOrderApproval(createdItem.ID, poAmountForApproval);

        // Send notification to the approver chosen by the threshold logic (same as form & PR/MR)
        if (webPartContext) {
          try {
            const updated = await sharePointService.getListData(newListName, `ID eq ${createdItem.ID}`, 1, 0, 'Project_Code,Vendor');
            if (updated.length > 0) {
              const fresh = updated[0];
              const ca = fresh?.Current_Approver;
              const nextApprover = ca
                ? (typeof ca === 'object' ? (ca.Email || ca.Title || ca.LoginName) : ca)
                : null;

              if (nextApprover) {
                const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
                const notificationService = new NotificationService(graphClient);
                const projectName = fresh.Project_Code?.Project_Code || fresh.Project_Code?.Title || 'Unknown Project';
                const deepLink = `${window.location.origin}${window.location.pathname}?po=${createdItem.ID}`;
                const title = fresh.PO_Number || fresh.Title || (newItemData as any).Title || createdItem.Title;

                await notificationService.sendMaterialRequestApprovalNotification({
                  requestTitle: title,
                  project: projectName,
                  amount: poAmountForApproval,
                  approverEmail: nextApprover,
                  deepLink,
                });
              }
            }
          } catch (notifyErr) {
            console.error('[Convert PO] Notification failed:', notifyErr);
          }
        }
      } else if (targetType === 'GRN' && createdItem && createdItem.ID) {
        // Best-effort post-creation steps for GRN (approval + Finance notification)
        // We never want a missing column on GRN or Approval Matrix to break the convert flow.
        try {
          await sharePointService.initializeGoodsReceivedNoteApproval(createdItem.ID);
        } catch (e) {
          console.warn('[Convert to GRN] Approval init skipped (non-fatal):', e);
        }

        try {
          if (record.ID) {
            await sharePointService.notifyFinanceManagerAfterGRN(createdItem.ID, record.ID);
          }
        } catch (e) {
          console.warn('[Convert to GRN] Finance notification skipped (non-fatal):', e);
        }
      }

      const targetLabel = targetType === 'PR' ? 'Purchase Requisition' :
        targetType === 'PO' ? 'Purchase Order' : 'Goods Received Note';

      setSuccessMessage(`Converted to ${targetLabel} successfully!`);

      if (onConvert) {
        onConvert(targetType);
      }

      if (onRefresh) {
        setTimeout(() => {
          onRefresh();
          onDismiss();
        }, 1500);
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to convert record';
      setError(errorMessage);
    } finally {
      setIsProcessing(false);
    }
  };

  // === PO Conversion Dialog helpers (for capturing Vendor + Delivery Date when converting PR → PO) ===
  const openPOConversionDialog = async (): Promise<void> => {
    setIsPOVendorDialogOpen(true);
    setPoDialogSelectedVendorId(undefined);
    setPoDialogDeliveryDate('');
    setIsLoadingVendorsForDialog(true);

    try {
      if (sharePointService) {
        const vendors = await sharePointService.getVendors();
        setPoDialogVendors(vendors);
      }
    } catch (e) {
      console.error('[PO Convert Dialog] Failed to load vendors:', e);
    } finally {
      setIsLoadingVendorsForDialog(false);
    }
  };

  const closePOConversionDialog = (): void => {
    setIsPOVendorDialogOpen(false);
  };

  const confirmPOConversionWithDetails = async (): Promise<void> => {
    if (!poDialogSelectedVendorId) {
      // Vendor is mandatory for a complete PO (for reporting)
      return;
    }
    closePOConversionDialog();

    // Proceed with creation, passing the user-selected values
    await handleConvert('PO', {
      vendorId: poDialogSelectedVendorId,
      deliveryDate: poDialogDeliveryDate || null,
    });
  };

  // Determine what conversion is available based on current record type and status
  const canConvert = (): boolean => {
    if (!onConvert && !sharePointService) return false;
    if (hasExistingConversion) return false;

    const status = record.Approval_Status?.toLowerCase() || record.Status?.toLowerCase();

    switch (recordType) {
      case 'MR':
        return (status === 'approved' || record.Approval_Status === 'Approved') && (!record.Approval_Level || record.Approval_Level >= 2);
      case 'PR':
        return (status === 'approved' || record.Approval_Status === 'Approved') && (!record.Approval_Level || record.Approval_Level >= 2);
      case 'PO':
        return status === 'approved' || record.Approval_Status === 'Approved';
      case 'GRN':
        return false;
      default:
        return false;
    }
  };

  const getConversionTarget = (): 'PR' | 'PO' | 'GRN' | null => {
    const status = record.Approval_Status?.toLowerCase() || record.Status?.toLowerCase();

    if (status === 'approved') {
      switch (recordType) {
        case 'MR':
          return 'PR';
        case 'PR':
          return 'PO';
        case 'PO':
          return 'GRN';
      }
    }
    return null;
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={getRecordTypeLabel()}
      closeButtonAriaLabel="Close"
      isLightDismiss
    >
      <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
        <div className={classNames.panelContent}>
          {/* Error/Success Messages */}
          {error && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {error}
              </MessageBar>
            </div>
          )}

          {successMessage && (
            <div style={{ marginBottom: '16px' }}>
              <MessageBar messageBarType={MessageBarType.success} isMultiline>
                {successMessage}
              </MessageBar>
            </div>
          )}

          {/* Header: Record ID + Status Badge */}
          <div className={classNames.section}>
            <div className={classNames.headerContainer}>
              <Text className={classNames.recordId}>ID: {record.ID || 'N/A'}</Text>
              {record.Status && (
                <span
                  className={classNames.statusBadge}
                  style={{
                    backgroundColor: getStatusBackgroundColor(record.Status),
                    color: getStatusTextColor(record.Status),
                  }}
                >
                  {record.Status}
                </span>
              )}
            </div>
          </div>

          <Separator />

          {/* Section 1: Record Details */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              Record Details
            </Text>

            {/* Record Number */}
            <div style={{ marginBottom: '16px' }}>
              <Label className={classNames.fieldLabel}>Record Number</Label>
              <Text className={classNames.fieldValue}>{record.Title || 'N/A'}</Text>
            </div>

            {/* Project Code */}
            {record.Project_Code && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Project Code</Label>
                <Text className={classNames.fieldValue}>{getDisplayValue(record.Project_Code)}</Text>
              </div>
            )}

            {/* Material (MR) */}
            {recordType === 'MR' && record.Material && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Material</Label>
                <Text className={classNames.fieldValue}>{getDisplayValue(record.Material)}</Text>
              </div>
            )}

            {/* Description (PR) */}
            {recordType === 'PR' && record.Description && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Description</Label>
                <Text className={classNames.fieldValue}>{record.Description}</Text>
              </div>
            )}

            {/* Vendor (PO, GRN) */}
            {(recordType === 'PO' || recordType === 'GRN') && record.Vendor && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Vendor</Label>
                <Text className={classNames.fieldValue}>{getDisplayValue(record.Vendor)}</Text>
              </div>
            )}

            {/* Quantity */}
            {record.Quantity && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Quantity</Label>
                <Text className={classNames.fieldValue}>
                  {record.Quantity} {record.UOM || 'PCS'}
                </Text>
              </div>
            )}

            {/* Unit Price (PO) */}
            {recordType === 'PO' && record.UnitPrice && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Unit Price</Label>
                <Text className={classNames.fieldValue}>
                  ${record.UnitPrice.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </Text>
              </div>
            )}

            {/* Total Amount (PO) */}
            {recordType === 'PO' && record.TotalAmount && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Total Amount</Label>
                <Text className={classNames.fieldValue} style={{ fontWeight: 600, color: theme.palette.themePrimary }}>
                  ${record.TotalAmount.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </Text>
              </div>
            )}

            {/* Estimated Cost (PR) */}
            {recordType === 'PR' && record.EstimatedCost && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Estimated Cost</Label>
                <Text className={classNames.fieldValue}>
                  ${record.EstimatedCost.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </Text>
              </div>
            )}

            {/* Approval Status */}
            {record.Approval_Status && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Approval Status</Label>
                <Text className={classNames.fieldValue} style={{ color: getStatusColor(record.Approval_Status), fontWeight: 600 }}>
                  {record.Approval_Status}
                </Text>
              </div>
            )}

            {/* Request Date */}
            {record.Request_Date && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Request Date</Label>
                <Text className={classNames.fieldValue}>
                  {new Date(record.Request_Date).toLocaleDateString()}
                </Text>
              </div>
            )}

            {/* Received Date (GRN) */}
            {recordType === 'GRN' && record.Received_Date && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Received Date</Label>
                <Text className={classNames.fieldValue}>
                  {new Date(record.Received_Date).toLocaleDateString()}
                </Text>
              </div>
            )}

            {/* Rejection Reason */}
            {record.Rejection_Reason && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Rejection Reason</Label>
                <Text className={classNames.fieldValue} style={{ color: theme.palette.red }}>
                  {record.Rejection_Reason}
                </Text>
              </div>
            )}

            {/* Notes */}
            {record.Notes && (
              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Notes</Label>
                <Text className={classNames.fieldValue}>{record.Notes}</Text>
              </div>
            )}
          </div>

          <Separator />

          {/* Section 2: QR Code */}
          <div className={classNames.section}>
            <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
              QR Code
            </Text>
            <div className={classNames.qrCodeContainer}>
              {qrCodeUrl && (
                <Image
                  src={qrCodeUrl}
                  alt={`QR Code for ${record.Title || record.ID}`}
                  className={classNames.qrImage}
                  width={150}
                  height={150}
                />
              )}
            </div>
            <Text variant="small" block style={{ textAlign: 'center', color: theme.palette.neutralSecondary }}>
              Scan to view record details
            </Text>
          </div>

          <Separator />

          {/* Section 3: Approval Actions */}
          {showApprovalButtons && (
            <div className={classNames.section}>
              <Text variant="large" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                Actions
              </Text>

              {/* Rejection Comment Box */}
              {showRejectComment && (
                <div className={classNames.rejectSection}>
                  <TextField
                    label="Rejection Reason"
                    placeholder="Please provide a reason for rejection"
                    multiline
                    rows={3}
                    value={rejectComment}
                    onChange={(_, value) => setRejectComment(value || '')}
                    disabled={isProcessing}
                  />
                  <div style={{ display: 'flex', gap: '8px', marginTop: '12px' }}>
                    <PrimaryButton
                      text="Confirm Rejection"
                      onClick={handleReject}
                      disabled={isProcessing || !rejectComment.trim()}
                      iconProps={{ iconName: 'Cancel' }}
                      style={{ backgroundColor: theme.palette.red, borderColor: theme.palette.red }}
                    />
                    <DefaultButton
                      text="Cancel"
                      onClick={() => {
                        setShowRejectComment(false);
                        setRejectComment('');
                      }}
                      disabled={isProcessing}
                    />
                  </div>
                </div>
              )}

               {/* Action Buttons */}
               {!showRejectComment && (
                 <div className={classNames.buttonGroup}>
                   {/* Approve Comment */}
                   {onApprove && canApprove && (record.Approval_Status === 'Pending' || record.Status === 'Draft' || record.Status === 'Pending Approval') && (
                     <TextField
                       label="Approval Comment (optional)"
                       placeholder="Add any notes for this approval"
                       multiline
                       rows={2}
                       value={approveComment}
                       onChange={(_, value) => setApproveComment(value || '')}
                       disabled={isProcessing}
                       styles={{ root: { width: '100%', marginBottom: '8px' } }}
                     />
                   )}
                   {/* Approve Button - only if user is authorized approver */}
                   {onApprove && canApprove && (record.Approval_Status === 'Pending' || record.Status === 'Draft' || record.Status === 'Pending Approval') && (
                     <PrimaryButton
                       text="Approve"
                      onClick={handleApprove}
                      disabled={isProcessing}
                      iconProps={{ iconName: 'CheckMark' }}
                      style={{ backgroundColor: theme.palette.green, borderColor: theme.palette.green }}
                    />
                  )}

                  {/* Reject Button - only if user is authorized approver */}
                  {onReject && canApprove && (record.Approval_Status === 'Pending' || record.Status === 'Draft' || record.Status === 'Pending Approval') && (
                    <DefaultButton
                      text="Reject"
                      onClick={() => setShowRejectComment(true)}
                      disabled={isProcessing}
                      iconProps={{ iconName: 'Cancel' }}
                      style={{ color: theme.palette.red, borderColor: theme.palette.red }}
                    />
                  )}

                  {/* Track Approval Button - always available for non-GRN */}
                  {onTrackApproval && recordType !== 'GRN' && (
                    <DefaultButton
                      text="Track Approval"
                      onClick={onTrackApproval}
                      disabled={isProcessing}
                      iconProps={{ iconName: 'CheckList' }}
                    />
                  )}

                   {/* PDF Preview & Download - only for Purchase Orders */}
                   {recordType === 'PO' && record.Title && (
                     <>
                       <DefaultButton
                         text="Preview PDF"
                         onClick={async () => {
                           // Build a robust path to the document library where the dynamically generated PDF is stored
                           const webUrl = pageContext?.web?.absoluteUrl || '';
                           const libraryPath = 'Shared Documents/Procurement/PO';
                           const safeTitle = encodeURIComponent(record.Title);
                           const pdfUrl = `${webUrl}/${libraryPath.replace(/ /g, '%20')}/${safeTitle}.pdf`;

                           // Best-effort: fetch linked PRs (multi + legacy) and trigger enriched PDF generation
                           let linkedPrIds: number[] = [];
                           if (sharePointService && record.ID) {
                             try {
                               const poData = await sharePointService.getListData(
                                 'PRC_Purchase_Order_Register',
                                 `ID eq ${record.ID}`,
                                 1,
                                 0,
                                 'PR_References,PR_NumberId'
                               );
                               if (poData.length > 0) {
                                 const po = poData[0];
                                 const legacy = getLookupId(po.PR_NumberId);
                                 if (legacy) linkedPrIds.push(legacy);
                                 const refs = po.PR_References || [];
                                 if (Array.isArray(refs)) {
                                   refs.forEach((r: any) => {
                                     const id = getLookupId(r);
                                     if (id) linkedPrIds.push(id);
                                   });
                                 }
                               }
                               await sharePointService.triggerPurchaseOrderPdfGeneration(record.Title, record.ID, linkedPrIds);
                             } catch (e) {
                               console.log('[PO] PDF generation trigger (non-blocking):', e);
                               await sharePointService.triggerPurchaseOrderPdfGeneration(record.Title, record.ID, []);
                             }
                           }

                           window.open(pdfUrl, '_blank');
                         }}
                         disabled={isProcessing}
                         iconProps={{ iconName: 'Preview' }}
                       />
                       <DefaultButton
                         text="Download PDF"
                         onClick={async () => {
                           const webUrl = pageContext?.web?.absoluteUrl || '';
                           const libraryPath = 'Shared Documents/Procurement/PO';
                           const safeTitle = encodeURIComponent(record.Title);
                           const pdfUrl = `${webUrl}/${libraryPath.replace(/ /g, '%20')}/${safeTitle}.pdf`;

                           let linkedPrIds: number[] = [];
                           if (sharePointService && record.ID) {
                             try {
                               const poData = await sharePointService.getListData(
                                 'PRC_Purchase_Order_Register',
                                 `ID eq ${record.ID}`,
                                 1,
                                 0,
                                 'PR_References,PR_NumberId'
                               );
                               if (poData.length > 0) {
                                 const po = poData[0];
                                 const legacy = getLookupId(po.PR_NumberId);
                                 if (legacy) linkedPrIds.push(legacy);
                                 const refs = po.PR_References || [];
                                 if (Array.isArray(refs)) {
                                   refs.forEach((r: any) => {
                                     const id = getLookupId(r);
                                     if (id) linkedPrIds.push(id);
                                   });
                                 }
                               }
                               await sharePointService.triggerPurchaseOrderPdfGeneration(record.Title, record.ID, linkedPrIds);
                             } catch (e) {
                               console.log('[PO] PDF generation trigger (non-blocking):', e);
                               await sharePointService.triggerPurchaseOrderPdfGeneration(record.Title, record.ID, []);
                             }
                           }

                           const link = document.createElement('a');
                           link.href = pdfUrl;
                           link.download = `${record.Title}.pdf`;
                           document.body.appendChild(link);
                           link.click();
                           document.body.removeChild(link);
                         }}
                         disabled={isProcessing}
                         iconProps={{ iconName: 'Download' }}
                       />
                     </>
                   )}

                   {/* Convert Button */}
                   {canConvert() && (
                     <DefaultButton
                        text={`Convert to ${getConversionTarget()}`}
                        onClick={() => {
                          const target = getConversionTarget();

                          // Special handling for PR → PO to ensure Vendor and Delivery Date are captured
                          // (critical for PO reporting). This follows ERP best practice.
                          if (target === 'PO' && recordType === 'PR') {
                            const hasVendor = !!(getLookupId(record.VendorId) || getLookupId(record.Vendor));
                            const hasDelivery = !!(record.Delivery_Date || record.Required_Date || record.DeliveryDate);

                            if (hasVendor && hasDelivery) {
                              // All required data present → silent direct conversion (fast path)
                              handleConvert('PO');
                            } else {
                              // Missing critical fields → open focused dialog to collect them
                              // without forcing the user through the full PO creation form.
                              openPOConversionDialog();
                            }
                          } else {
                            handleConvert(target!);
                          }
                        }}
                       disabled={isProcessing}
                       iconProps={{ iconName: 'ArrowRight' }}
                       style={{ backgroundColor: theme.palette.themeLight, borderColor: theme.palette.themeLight }}
                     />
                   )}
                </div>
              )}

              {isProcessing && (
                <Text variant="small" style={{ marginTop: '12px', color: theme.palette.neutralSecondary }}>
                  Processing...
                </Text>
              )}
            </div>
          )}

          {/* Focused PO Conversion Dialog - captures Vendor + Delivery Date when converting PR → PO */}
          {/* This ensures complete data for PO reporting without opening the full PO form (ERP best practice) */}
          <Dialog
            hidden={!isPOVendorDialogOpen}
            onDismiss={closePOConversionDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Complete Purchase Order Details',
              subText: 'Please select the Vendor and optional Delivery Date to create a complete PO ready for reporting.',
            }}
            modalProps={{ isBlocking: true }}
          >
            {isLoadingVendorsForDialog && (
              <div style={{ padding: '12px 0', textAlign: 'center' }}>
                <Spinner size={SpinnerSize.medium} label="Loading vendors..." />
              </div>
            )}

            {!isLoadingVendorsForDialog && (
              <>
                <div style={{ marginBottom: 16 }}>
                  <Dropdown
                    label="Vendor"
                    required
                    options={poDialogVendors.map((v: any) => ({
                      key: v.Id || v.ID,
                      text: v.Vendor_Name || v.Title || `Vendor ${(v.Id || v.ID)}`,
                    }))}
                    selectedKey={poDialogSelectedVendorId}
                    onChange={(_, option) => setPoDialogSelectedVendorId(option ? Number(option.key) : undefined)}
                    placeholder="Select Vendor (required)"
                  />
                </div>

                <div style={{ marginBottom: 8 }}>
                  <TextField
                    label="Delivery Date"
                    type="date"
                    value={poDialogDeliveryDate}
                    onChange={(_, value) => setPoDialogDeliveryDate(value || '')}
                  />
                </div>
              </>
            )}

            <DialogFooter>
              <PrimaryButton
                onClick={confirmPOConversionWithDetails}
                text="Create Purchase Order"
                disabled={!poDialogSelectedVendorId || isLoadingVendorsForDialog}
              />
              <DefaultButton onClick={closePOConversionDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </ScrollablePane>
    </Panel>
  );
};

export default ProcurementDetailsPanel;
