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
import type { IWarehouse } from './WarehousesModule';
import WarehouseForm from './WarehouseForm';
import { SharePointService } from '../services/SharePointService';
import { SHAREPOINT_LISTS } from '../constants/SharePointListNames';

export interface IWarehouseDetailsPanelProps {
  isOpen: boolean;
  warehouse: IWarehouse | undefined;
  onDismiss: () => void;
  onRefresh?: () => void;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

const WarehouseDetailsPanel: React.FC<IWarehouseDetailsPanelProps> = ({
  isOpen,
  warehouse,
  onDismiss,
  onRefresh,
  spHttpClient,
  pageContext,
}) => {
  const theme = getTheme();
  const [isEditMode, setIsEditMode] = React.useState(false);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState(false);
  const [isDeleting, setIsDeleting] = React.useState(false);

  const classNames = mergeStyleSets({
    panelContent: { padding: '20px', paddingTop: '32px' },
    section: { marginBottom: '24px' },
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
    actionButton: { flex: 1, minWidth: '120px' },
  });

  if (!warehouse) {
    return null;
  }

  if (isEditMode) {
    return (
      <WarehouseForm
        isOpen={true}
        onDismiss={() => setIsEditMode(false)}
        onSubmitSuccess={() => {
          setIsEditMode(false);
          onRefresh?.();
          onDismiss();
        }}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        editMode={true}
        editWarehouse={warehouse}
      />
    );
  }

  const getStatusColor = (status: string): string => {
    return status === 'Active' ? theme.palette.green : theme.palette.red;
  };

  const handleDelete = async (): Promise<void> => {
    setIsDeleting(true);
    try {
      const sharePointService = new SharePointService(spHttpClient, pageContext);
      await sharePointService.deleteListItem(SHAREPOINT_LISTS.WAREHOUSES_MASTER, warehouse.ID);
      setIsDeleteDialogOpen(false);
      onRefresh?.();
      onDismiss();
    } catch (error) {
      console.error('Error deleting warehouse:', error);
    } finally {
      setIsDeleting(false);
    }
  };

  return (
    <>
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText="Warehouse Details"
        closeButtonAriaLabel="Close"
        isLightDismiss
      >
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <div className={classNames.panelContent}>
            <div className={classNames.section} style={{ marginTop: '16px' }}>
              <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                {warehouse.Warehouse_Name}
              </Text>
              <div
                className={classNames.statusBadge}
                style={{
                  backgroundColor: getStatusColor(warehouse.Status),
                  color: theme.palette.white,
                }}
              >
                {warehouse.Status}
              </div>
            </div>

            <Separator />

            <div className={classNames.section}>
              <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
                Warehouse Information
              </Text>
              <Label className={classNames.fieldLabel}>Warehouse Code</Label>
              <Text className={classNames.fieldValue}>{warehouse.WarehouseCode}</Text>
              <Label className={classNames.fieldLabel}>Warehouse Name</Label>
              <Text className={classNames.fieldValue}>{warehouse.Warehouse_Name}</Text>
              <Label className={classNames.fieldLabel}>Physical Location</Label>
              <Text className={classNames.fieldValue}>{warehouse.Location || 'N/A'}</Text>
              <Label className={classNames.fieldLabel}>ID</Label>
              <Text className={classNames.fieldValue}>{warehouse.ID}</Text>
            </div>

            <Separator />

            <div className={classNames.section}>
              <div className={classNames.buttonGroup}>
                <PrimaryButton
                  text="Edit"
                  onClick={() => setIsEditMode(true)}
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

      <Dialog
        hidden={!isDeleteDialogOpen}
        onDismiss={() => setIsDeleteDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Warehouse',
          subText: `Are you sure you want to delete "${warehouse.Warehouse_Name}"? This action cannot be undone.`,
        }}
        modalProps={{ isBlocking: true, styles: { main: { maxWidth: 450 } } }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDelete}
            text="Delete"
            disabled={isDeleting}
            iconProps={{ iconName: 'Delete' }}
            styles={{ root: { backgroundColor: theme.palette.red, color: theme.palette.white } }}
          />
          <DefaultButton onClick={() => setIsDeleteDialogOpen(false)} text="Cancel" disabled={isDeleting} />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default WarehouseDetailsPanel;
