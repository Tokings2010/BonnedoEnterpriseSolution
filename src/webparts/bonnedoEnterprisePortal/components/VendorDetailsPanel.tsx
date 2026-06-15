import * as React from 'react';
import {
  Panel,
  PanelType,
  Text,
  PrimaryButton,
  DefaultButton,
  Separator,
  Label,
  getTheme,
  mergeStyleSets,
  ScrollablePane,
  ScrollbarVisibility,
  Dialog,
  DialogType,
  DialogFooter,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IVendor } from './VendorsModule';
import VendorForm from './VendorForm';
import { SharePointService } from '../services/SharePointService';

export interface IVendorDetailsPanelProps {
  isOpen: boolean;
  vendor: IVendor | undefined;
  onDismiss: () => void;
  onRefresh?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

const VendorDetailsPanel: React.FC<IVendorDetailsPanelProps> = ({
  isOpen,
  vendor,
  onDismiss,
  onRefresh,
  spHttpClient,
  pageContext,
}) => {
  const theme = getTheme();
  const [isEditMode, setIsEditMode] = React.useState(false);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState(false);
  const [isDeleting, setIsDeleting] = React.useState(false);

  const handleEdit = (): void => {
    setIsEditMode(true);
  };

  const handleEditFormDismiss = (): void => {
    setIsEditMode(false);
  };

  const handleEditFormSuccess = (): void => {
    setIsEditMode(false);
    if (onRefresh) {
      onRefresh();
    }
    onDismiss();
  };

  const handleDelete = async (): Promise<void> => {
    if (!vendor) return;

    setIsDeleting(true);
    try {
      const sharePointService = new SharePointService(spHttpClient, pageContext);
      await sharePointService.deleteListItem('ENT_Vendors_Master', vendor.ID);
      setIsDeleteDialogOpen(false);
      if (onRefresh) {
        onRefresh();
      }
      onDismiss();
    } catch (error) {
      console.error('Error deleting vendor:', error);
      // You might want to show an error message here
    } finally {
      setIsDeleting(false);
    }
  };

  const classNames = mergeStyleSets({
    panelContent: {
      padding: '20px',
      paddingTop: '32px',
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
    statusBadge: {
      display: 'inline-block',
      padding: '6px 12px',
      borderRadius: '4px',
      fontSize: '13px',
      fontWeight: 600,
      marginBottom: '16px',
    },
    buttonGroup: {
      display: 'flex',
      gap: '12px',
      marginTop: '24px',
      flexWrap: 'wrap',
    },
    actionButton: {
      flex: 1,
      minWidth: '120px',
    },
  });

  if (!vendor) {
    return null;
  }

  // Show edit form when in edit mode
  if (isEditMode) {
    return (
      <VendorForm
        isOpen={true}
        onDismiss={handleEditFormDismiss}
        onSubmitSuccess={handleEditFormSuccess}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        editMode={true}
        editVendor={{
          ID: vendor.ID,
          Vendor_Code: vendor.Vendor_Code,
          Vendor_Name: vendor.Vendor_Name,
          Vendor_Category: vendor.Vendor_Category,
          Vendor_Status: vendor.Vendor_Status,
        }}
      />
    );
  }

  const getStatusColor = (status: string): string => {
    switch (status?.toLowerCase()) {
      case 'active':
        return theme.palette.green;
      case 'inactive':
        return theme.palette.red;
      case 'pending':
        return theme.palette.orange;
      case 'blocked':
        return theme.palette.red;
      default:
        return theme.palette.neutralSecondary;
    }
  };

  return (
    <>
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText="Vendor Details"
        closeButtonAriaLabel="Close"
        isLightDismiss
      >
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <div className={classNames.panelContent}>
            {/* Vendor Header */}
            <div className={classNames.section} style={{ marginTop: '16px' }}>
              <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                {vendor.Vendor_Name}
              </Text>
              <div
                className={classNames.statusBadge}
                style={{
                  backgroundColor: getStatusColor(vendor.Vendor_Status),
                  color: theme.palette.white,
                }}
              >
                {vendor.Vendor_Status}
              </div>
            </div>

            <Separator />

            {/* Vendor Information */}
            <div className={classNames.section}>
              <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
                Vendor Information
              </Text>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Vendor Code</Label>
                <Text className={classNames.fieldValue}>{vendor.Vendor_Code}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Vendor Name</Label>
                <Text className={classNames.fieldValue}>{vendor.Vendor_Name}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Category</Label>
                <Text className={classNames.fieldValue}>{vendor.Vendor_Category || 'N/A'}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Status</Label>
                <Text className={classNames.fieldValue} style={{ color: getStatusColor(vendor.Vendor_Status) }}>
                  {vendor.Vendor_Status}
                </Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>ID</Label>
                <Text className={classNames.fieldValue}>{vendor.ID}</Text>
              </div>
            </div>

            <Separator />

            {/* Action Buttons */}
            <div className={classNames.section}>
              <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
                Actions
              </Text>

              <div className={classNames.buttonGroup}>
                <PrimaryButton
                  text="Edit"
                  onClick={handleEdit}
                  iconProps={{ iconName: 'Edit' }}
                  className={classNames.actionButton}
                />
                <DefaultButton
                  text="Delete"
                  onClick={() => setIsDeleteDialogOpen(true)}
                  iconProps={{ iconName: 'Delete' }}
                  className={classNames.actionButton}
                  styles={{ root: { backgroundColor: theme.palette.red, color: theme.palette.white } }}
                />
              </div>
            </div>
          </div>
        </ScrollablePane>
      </Panel>

      {/* Delete Confirmation Dialog */}
      <Dialog
        hidden={!isDeleteDialogOpen}
        onDismiss={() => setIsDeleteDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Vendor',
          subText: `Are you sure you want to delete "${vendor.Vendor_Name}"? This action cannot be undone.`,
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 450 } },
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDelete}
            text="Delete"
            disabled={isDeleting}
            iconProps={{ iconName: 'Delete' }}
            styles={{ root: { backgroundColor: theme.palette.red, color: theme.palette.white } }}
          />
          <DefaultButton
            onClick={() => setIsDeleteDialogOpen(false)}
            text="Cancel"
            disabled={isDeleting}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default VendorDetailsPanel;