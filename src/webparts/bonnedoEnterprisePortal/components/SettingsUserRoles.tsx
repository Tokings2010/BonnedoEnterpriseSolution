import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    IconButton,
    PrimaryButton,
    DefaultButton,
    Dialog,
    DialogType,
    DialogFooter,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';
import { IListItem, SharePointService } from '../services/SharePointService';
import { PermissionService } from '../services/PermissionService';
import UserRoleForm from './UserRoleForm';
import UserRoleDetailsPanel from './UserRoleDetailsPanel';

// Settings User Roles Component
interface ISettingsUserRolesProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    onNewUserRole?: () => void;
    isMobileView?: boolean;
}

const SettingsUserRoles: React.FC<ISettingsUserRolesProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    onNewUserRole,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [refreshKey, setRefreshKey] = React.useState(0);

    // Delete dialog state
    const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState(false);
    const [selectedUserForDelete, setSelectedUserForDelete] = React.useState<IListItem | null>(null);
    const [isDeleting, setIsDeleting] = React.useState(false);

    // Details panel state (preview)
    const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
    const [selectedUserRole, setSelectedUserRole] = React.useState<IListItem | null>(null);

    // Edit form panel state
    const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);

    // Permission service for SharePoint group management
    const permissionService = React.useMemo(
        () => new PermissionService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    // SharePoint service for list operations
    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    const userRolesColumns: IDataGridColumn[] = [
        {
            key: 'User',
            name: 'User',
            fieldName: 'User',
            minWidth: 180,
            isResizable: true,
            onRender: (item: IListItem) => {
                // Handle User as lookup object
                const userName = item.User?.Title || item.User?.Name || item.User?.displayName || 'Unknown User';
                return <Text>{userName}</Text>;
            }
        },
        { key: 'Role', name: 'Role', fieldName: 'Role', minWidth: 150, isResizable: true },
        { key: 'Department', name: 'Department', fieldName: 'Department', minWidth: 150, isResizable: true },
        { key: 'Permissions', name: 'Permissions', fieldName: 'Permissions', minWidth: 250, isResizable: true },
        {
            key: 'Actions',
            name: 'Actions',
            fieldName: 'Actions',
            minWidth: 80,
            maxWidth: 80,
            isResizable: false,
            onRender: (item: IListItem) => {
                return (
                    <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Revoke Access"
                        ariaLabel="Revoke Access"
                        onClick={(e) => {
                            e.stopPropagation();
                            handleDeleteClick(item);
                        }}
                        styles={{
                            root: {
                                color: theme.palette.red,
                            }
                        }}
                    />
                );
            }
        },
    ];

    const handleRefresh = (): void => {
        setRefreshKey(prev => prev + 1);
        onRefresh?.();
    };

    // Handle row selection (single click) - opens details panel
    const handleRowSelected = (record: IListItem): void => {
        console.log('User role selected for preview:', record);
        setSelectedUserRole(record);
        setIsDetailsPanelOpen(true);
    };

    // Handle row double click - opens edit form directly
    const handleRowDoubleClick = (record: IListItem): void => {
        console.log('User role selected for edit:', record);
        setSelectedUserRole(record);
        setIsFormPanelOpen(true);
    };

    // Handle delete button click
    const handleDeleteClick = (record: IListItem): void => {
        setSelectedUserForDelete(record);
        setIsDeleteDialogOpen(true);
    };

    // Confirm delete
    const handleDeleteConfirm = async (): Promise<void> => {
        if (!selectedUserForDelete) return;

        setIsDeleting(true);
        try {
            // Get the user ID from the record
            const userId = selectedUserForDelete.UserId || selectedUserForDelete.User?.ID;

            // Remove user from SharePoint groups
            if (userId) {
                await permissionService.removeUserFromSite(userId);
            }

            // Note: Deleting the list item would require additional implementation
            // For now, we just remove from SharePoint groups

            setIsDeleteDialogOpen(false);
            setSelectedUserForDelete(null);
            setRefreshKey(prev => prev + 1);
            onRefresh?.();
        } catch (error) {
            console.error('Error deleting user role:', error);
        } finally {
            setIsDeleting(false);
        }
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px' },
        header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' },
        description: {
            color: theme.palette.neutralSecondary,
            marginBottom: '16px',
        },
    });

    return (
        <div className={classNames.container}>
            {/* Tab Header */}
            <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        User Roles
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage user roles and permissions
                    </Text>
                </div>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <PrimaryButton
                        text="+ New User Role"
                        onClick={onNewUserRole}
                        iconProps={{ iconName: 'Add' }}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={handleRefresh}
                        title="Refresh data"
                        ariaLabel="Refresh data"
                    />
                </div>
            </div>

            <Text variant="medium" className={classNames.description}>
                Manage user roles and their permissions in the system. Double-click a row to edit.
            </Text>

            <div style={{ height: 'calc(100vh - 280px)', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                <EnhancedDataGrid
                    key={refreshKey}
                    listName="SYS_User_Roles"
                    columns={userRolesColumns}
                    pageSize={10}
                    spHttpClient={spHttpClient}
                    pageContext={pageContext}
                    expandQuery="User"
                    onRowSelected={handleRowSelected}
                    onRowDoubleClick={handleRowDoubleClick}
                    showExport
                />
            </div>

            {/* Delete Confirmation Dialog */}
            <Dialog
                isOpen={isDeleteDialogOpen}
                onDismiss={() => setIsDeleteDialogOpen(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Revoke User Access',
                    subText: `Are you sure you want to revoke access for ${selectedUserForDelete?.User?.Title || 'this user'}? This will remove them from the SharePoint site groups.`,
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 450 } },
                }}
            >
                <DialogFooter>
                    <DefaultButton onClick={() => setIsDeleteDialogOpen(false)} text="Cancel" />
                    <PrimaryButton
                        onClick={handleDeleteConfirm}
                        text={isDeleting ? 'Removing...' : 'Revoke Access'}
                        disabled={isDeleting}
                    />
                </DialogFooter>
            </Dialog>

            {/* User Role Details Panel (Preview) */}
            <UserRoleDetailsPanel
                isOpen={isDetailsPanelOpen}
                userRole={selectedUserRole ? {
                    ID: selectedUserRole.ID,
                    User: selectedUserRole.User as { Title: string; ID: number },
                    Role: selectedUserRole.Role || '',
                    Department: selectedUserRole.Department || '',
                    Permissions: selectedUserRole.Permissions || '',
                } : undefined}
                onDismiss={() => setIsDetailsPanelOpen(false)}
                onRefresh={handleRefresh}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
            />

            {/* Edit User Role Form Panel */}
            <UserRoleForm
                isOpen={isFormPanelOpen}
                onDismiss={() => setIsFormPanelOpen(false)}
                onSubmitSuccess={() => {
                    setRefreshKey(prev => prev + 1);
                    onRefresh?.();
                }}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                editMode={!!selectedUserRole}
                editUserRole={selectedUserRole ? {
                    ID: selectedUserRole.ID,
                    User: selectedUserRole.User as { Title: string; ID: number },
                    Role: selectedUserRole.Role || '',
                    Department: selectedUserRole.Department || '',
                    Permissions: selectedUserRole.Permissions || '',
                } : undefined}
            />
        </div >
    );
};

export default SettingsUserRoles;
