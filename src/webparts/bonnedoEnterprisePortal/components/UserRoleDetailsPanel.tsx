import * as React from 'react';
import {
    Panel,
    PanelType,
    Text,
    PrimaryButton,
    DefaultButton,
    Separator,
    getTheme,
    mergeStyleSets,
    ScrollablePane,
    ScrollbarVisibility,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import UserRoleForm from './UserRoleForm';

export interface IUserRoleDetailsPanelProps {
    isOpen: boolean;
    userRole: {
        ID: number;
        User: { Title: string; ID: number };
        Role: string;
        Department: string;
        Permissions: string;
    } | undefined;
    onDismiss: () => void;
    onRefresh?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

const UserRoleDetailsPanel: React.FC<IUserRoleDetailsPanelProps> = ({
    isOpen,
    userRole,
    onDismiss,
    onRefresh,
    spHttpClient,
    pageContext,
    webPartContext,
}) => {
    const theme = getTheme();
    const [isEditMode, setIsEditMode] = React.useState(false);

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

    const classNames = mergeStyleSets({
        panelContent: {
            padding: '20px',
            paddingTop: '32px',
        },
        section: {
            marginBottom: '24px',
        },
        fieldLabel: {
            fontSize: '12px',
            fontWeight: 600,
            color: theme.palette.neutralSecondary,
            marginBottom: '4px',
            textTransform: 'uppercase',
        },
        fieldValue: {
            fontSize: '14px',
            color: theme.palette.neutralPrimary,
            marginBottom: '16px',
        },
        buttonGroup: {
            display: 'flex',
            gap: '8px',
            flexWrap: 'wrap',
        },
        actionButton: {
            minWidth: '100px',
        },
    });

    if (!userRole) {
        return null;
    }

    // If in edit mode, render the form instead
    if (isEditMode) {
        return (
            <UserRoleForm
                isOpen={true}
                onDismiss={handleEditFormDismiss}
                onSubmitSuccess={handleEditFormSuccess}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                webPartContext={webPartContext}
                editMode={true}
                editUserRole={userRole}
            />
        );
    }

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            headerText="User Role Details"
            closeButtonAriaLabel="Close"
            isLightDismiss
        >
            <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                <div className={classNames.panelContent}>
                    {/* User Header */}
                    <div className={classNames.section} style={{ marginTop: '16px' }}>
                        <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '12px' }}>
                            {userRole.User?.Title || 'Unknown User'}
                        </Text>
                    </div>

                    <Separator />

                    {/* User Role Information */}
                    <div className={classNames.section}>
                        <Text variant="large" block style={{ fontWeight: 600, marginBottom: '16px' }}>
                            Role Information
                        </Text>

                        {/* Role */}
                        <div>
                            <div className={classNames.fieldLabel}>Role</div>
                            <div className={classNames.fieldValue}>{userRole.Role || 'Not assigned'}</div>
                        </div>

                        {/* Department */}
                        <div>
                            <div className={classNames.fieldLabel}>Department</div>
                            <div className={classNames.fieldValue}>{userRole.Department || 'Not assigned'}</div>
                        </div>

                        {/* Permissions */}
                        <div>
                            <div className={classNames.fieldLabel}>Permissions</div>
                            <div className={classNames.fieldValue} style={{ whiteSpace: 'pre-wrap' }}>
                                {userRole.Permissions || 'No permissions specified'}
                            </div>
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
                        </div>
                    </div>
                </div>
            </ScrollablePane>
        </Panel>
    );
};

export default UserRoleDetailsPanel;
