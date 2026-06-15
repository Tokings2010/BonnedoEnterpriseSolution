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
    MessageBar,
    MessageBarType,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ApprovalMatrixForm from './ApprovalMatrixForm';

export interface IApprovalMatrixDetailsPanelProps {
    isOpen: boolean;
    approvalMatrix: {
        ID: number;
        Module: string;
        Min_Amount: number;
        Max_Amount: number;
        Stage?: number;
        Approver_Role: string;
        Approver_UserId: number;
        Approver_User?: { Title: string; ID: number };
    } | undefined;
    onDismiss: () => void;
    onRefresh?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

const ApprovalMatrixDetailsPanel: React.FC<IApprovalMatrixDetailsPanelProps> = ({
    isOpen,
    approvalMatrix,
    onDismiss,
    onRefresh,
    spHttpClient,
    pageContext,
    webPartContext,
}) => {
    const theme = getTheme();
    const [isEditMode, setIsEditMode] = React.useState(false);
    const [isDeleting, setIsDeleting] = React.useState(false);
    const [deleteMessage, setDeleteMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);

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
        if (!approvalMatrix) return;

        const confirmed = window.confirm('Are you sure you want to delete this approval rule?');
        if (!confirmed) return;

        setIsDeleting(true);
        setDeleteMessage(null);

        try {
            const webUrl = pageContext.web.absoluteUrl;
            const deleteUrl = `${webUrl}/_api/web/lists/getByTitle('SYS_Approval_Matrix')/items(${approvalMatrix.ID})`;

            const response = await spHttpClient.post(
                deleteUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'DELETE',
                        'If-Match': '*'
                    }
                }
            );

            if (response.ok) {
                setDeleteMessage({
                    type: MessageBarType.success,
                    text: 'Approval rule deleted successfully!'
                });
                setTimeout(() => {
                    if (onRefresh) {
                        onRefresh();
                    }
                    onDismiss();
                }, 1500);
            } else {
                throw new Error('Failed to delete approval rule');
            }
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to delete approval rule';
            setDeleteMessage({
                type: MessageBarType.error,
                text: errorMessage
            });
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
            fontSize: '12px',
            fontWeight: 600,
            color: theme.palette.neutralSecondary,
            marginBottom: '4px',
            textTransform: 'uppercase' as const,
        },
        fieldValue: {
            fontSize: '14px',
            color: theme.palette.neutralPrimary,
        },
        buttonGroup: {
            display: 'flex',
            gap: '8px',
            marginTop: '24px',
        },
        deleteSection: {
            marginTop: '32px',
            paddingTop: '16px',
            borderTop: `1px solid ${theme.palette.neutralLight}`,
        },
    });

    if (isEditMode && approvalMatrix) {
        return (
            <ApprovalMatrixForm
                isOpen={isOpen}
                onDismiss={handleEditFormDismiss}
                onSubmitSuccess={handleEditFormSuccess}
                spHttpClient={spHttpClient}
                pageContext={pageContext}
                webPartContext={webPartContext}
                editMode={true}
                editRecord={{
                    ID: approvalMatrix.ID,
                    Module: approvalMatrix.Module,
                    Min_Amount: approvalMatrix.Min_Amount,
                    Max_Amount: approvalMatrix.Max_Amount,
                    Stage: approvalMatrix.Stage,
                    Approver_Role: approvalMatrix.Approver_Role,
                    Approver_UserId: approvalMatrix.Approver_UserId,
                    Approver_User: approvalMatrix.Approver_User
                }}
            />
        );
    }

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            headerText="Approval Rule Details"
            closeButtonAriaLabel="Close"
        >
            <div className={classNames.panelContent}>
                {approvalMatrix && (
                    <div style={{ display: 'flex', gap: '8px', marginBottom: '16px', justifyContent: 'flex-end' }}>
                        <PrimaryButton
                            text="Edit"
                            iconProps={{ iconName: 'Edit' }}
                            onClick={handleEdit}
                        />
                    </div>
                )}

                {deleteMessage && (
                    <div style={{ marginBottom: '16px' }}>
                        <MessageBar messageBarType={deleteMessage.type} onDismiss={() => setDeleteMessage(null)}>
                            {deleteMessage.text}
                        </MessageBar>
                    </div>
                )}

                {approvalMatrix ? (
                    <>
                        <div className={classNames.section}>
                            <div className={classNames.fieldLabel}>Module</div>
                            <div className={classNames.fieldValue}>{approvalMatrix.Module || '-'}</div>
                        </div>

                        <Separator />

                        <div className={classNames.section}>
                            <div className={classNames.fieldLabel}>Amount Range</div>
                            <div className={classNames.fieldValue}>
                                ₦{approvalMatrix.Min_Amount?.toLocaleString() || '0'} - ₦{approvalMatrix.Max_Amount?.toLocaleString() || '0'}
                            </div>
                        </div>

                        <Separator />

                        <div className={classNames.section}>
                            <div className={classNames.fieldLabel}>Stage</div>
                            <div className={classNames.fieldValue}>{approvalMatrix.Stage ?? '-'}</div>
                        </div>

                        <Separator />

                        <div className={classNames.section}>
                            <div className={classNames.fieldLabel}>Approver Role</div>
                            <div className={classNames.fieldValue}>{approvalMatrix.Approver_Role || '-'}</div>
                        </div>

                        <Separator />

                        <div className={classNames.section}>
                            <div className={classNames.fieldLabel}>Approver User</div>
                            <div className={classNames.fieldValue}>
                                {approvalMatrix.Approver_User?.Title ||
                                    (approvalMatrix.Approver_UserId ? `User ID: ${approvalMatrix.Approver_UserId}` : '-')}
                            </div>
                        </div>

                        <div className={classNames.deleteSection}>
                            <Text variant="medium" style={{ color: theme.palette.red }}>
                                Danger Zone
                            </Text>
                            <div className={classNames.buttonGroup}>
                                <DefaultButton
                                    text="Delete Approval Rule"
                                    iconProps={{ iconName: 'Delete' }}
                                    onClick={handleDelete}
                                    disabled={isDeleting}
                                    style={{
                                        backgroundColor: '#fff0f0',
                                        color: '#d13438'
                                    }}
                                />
                            </div>
                        </div>
                    </>
                ) : (
                    <Text>No approval rule selected</Text>
                )}
            </div>
        </Panel>
    );
};

export default ApprovalMatrixDetailsPanel;
