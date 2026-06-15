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
import { IMaterial } from './MaterialsModule';
import MaterialForm from './MaterialForm';
import { SharePointService } from '../services/SharePointService';

export interface IMaterialDetailsPanelProps {
  isOpen: boolean;
  material: IMaterial | undefined;
  onDismiss: () => void;
  onRefresh?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

const MaterialDetailsPanel: React.FC<IMaterialDetailsPanelProps> = ({
  isOpen,
  material,
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
    if (!material) return;

    setIsDeleting(true);
    try {
      const sharePointService = new SharePointService(spHttpClient, pageContext);
      await sharePointService.deleteListItem('ENT_Materials_Master', material.ID);
      setIsDeleteDialogOpen(false);
      if (onRefresh) {
        onRefresh();
      }
      onDismiss();
    } catch (error) {
      console.error('Error deleting material:', error);
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

  if (!material) {
    return null;
  }

  // Show edit form when in edit mode
  if (isEditMode) {
    return (
      <MaterialForm
        isOpen={true}
        onDismiss={handleEditFormDismiss}
        onSubmitSuccess={handleEditFormSuccess}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        editMode={true}
        editMaterial={{
          ID: material.ID,
          Material_Code: material.Material_Code,
          Material_Name: material.Material_Name,
          Category: material.Category,
          UOM: material.UOM,
          Standard_Cost: material.Standard_Cost,
          Active: material.Active,
        }}
      />
    );
  }

  const getStatusColor = (active: boolean): string => {
    return active ? theme.palette.green : theme.palette.red;
  };

  const getStatusText = (active: boolean): string => {
    return active ? 'Active' : 'Inactive';
  };

  return (
    <>
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText="Material Details"
        closeButtonAriaLabel="Close"
        isLightDismiss
      >
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <div className={classNames.panelContent}>
            {/* Material Header */}
            <div className={classNames.section} style={{ marginTop: '16px' }}>
              <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                {material.Material_Name}
              </Text>
              <div
                className={classNames.statusBadge}
                style={{
                  backgroundColor: getStatusColor(material.Active),
                  color: theme.palette.white,
                }}
              >
                {getStatusText(material.Active)}
              </div>
            </div>

            <Separator />

            {/* Material Information */}
            <div className={classNames.section}>
              <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
                Material Information
              </Text>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Material Code</Label>
                <Text className={classNames.fieldValue}>{material.Material_Code}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Material Name</Label>
                <Text className={classNames.fieldValue}>{material.Material_Name}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Category</Label>
                <Text className={classNames.fieldValue}>{material.Category || 'N/A'}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Unit of Measure</Label>
                <Text className={classNames.fieldValue}>{material.UOM || 'N/A'}</Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Standard Cost</Label>
                <Text className={classNames.fieldValue}>
                  ₦{material.Standard_Cost?.toLocaleString('en-NG', {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2,
                  }) || '0.00'}
                </Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>Status</Label>
                <Text className={classNames.fieldValue} style={{ color: getStatusColor(material.Active) }}>
                  {getStatusText(material.Active)}
                </Text>
              </div>

              <div style={{ marginBottom: '16px' }}>
                <Label className={classNames.fieldLabel}>ID</Label>
                <Text className={classNames.fieldValue}>{material.ID}</Text>
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
          title: 'Delete Material',
          subText: `Are you sure you want to delete "${material.Material_Name}"? This action cannot be undone.`,
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

export default MaterialDetailsPanel;