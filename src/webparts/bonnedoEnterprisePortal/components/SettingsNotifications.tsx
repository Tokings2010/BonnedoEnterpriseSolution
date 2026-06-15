import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    Toggle,
    PrimaryButton,
    IconButton,
    DefaultButton,
    MessageBar,
    MessageBarType,
    Label,
    Slider,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';

// Settings Notifications Component
interface ISettingsNotificationsProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const SettingsNotifications: React.FC<ISettingsNotificationsProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [settings, setSettings] = React.useState({
        EmailNotificationEnabled: true,
        TeamsNotificationEnabled: false,
        ApprovalReminderDays: 3,
    });
    const [loading, setLoading] = React.useState(true);
    const [saving, setSaving] = React.useState(false);
    const [saveMessage, setSaveMessage] = React.useState<string | null>(null);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    React.useEffect(() => {
        const fetchSettings = async () => {
            setLoading(true);
            try {
                const data = await sharePointService.getListData('SYS_Notification_Settings', undefined, 10);
                if (data.length > 0) {
                    const config = data[0];
                    setSettings({
                        EmailNotificationEnabled: config.EmailNotificationEnabled || true,
                        TeamsNotificationEnabled: config.TeamsNotificationEnabled || false,
                        ApprovalReminderDays: config.ApprovalReminderDays || 3,
                    });
                }
            } catch (error) {
                console.error('Error fetching notification settings:', error);
            }
            setLoading(false);
        };
        fetchSettings();
    }, [sharePointService]);

    const handleSave = async (): Promise<void> => {
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            const existing = await sharePointService.getListData('SYS_Notification_Settings', undefined, 1);

            if (existing.length > 0) {
                const url = `${webUrl}/_api/web/lists/getByTitle('SYS_Notification_Settings')/items(${existing[0].ID})`;
                await spHttpClient.post(url, SPHttpClient.configurations.v1, {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'MERGE',
                        'If-Match': '*',
                    },
                    body: JSON.stringify(settings),
                });
            } else {
                await sharePointService.createListItem('SYS_Notification_Settings', settings);
            }

            setSaveMessage('Notification settings saved successfully!');
            setTimeout(() => setSaveMessage(null), 3000);
        } catch (error) {
            setSaveMessage('Failed to save notification settings');
        }
        setSaving(false);
    };

    const handleReset = (): void => {
        setSettings({
            EmailNotificationEnabled: true,
            TeamsNotificationEnabled: false,
            ApprovalReminderDays: 3,
        });
    };

    const classNames = mergeStyleSets({
        container: { padding: '20px' },
        section: {
            marginBottom: '32px',
            padding: '24px',
            backgroundColor: theme.palette.white,
            borderRadius: '8px',
            border: `1px solid ${theme.palette.neutralLight}`,
        },
        sectionTitle: {
            fontSize: '18px',
            fontWeight: 600,
            marginBottom: '20px',
            color: theme.palette.neutralPrimary,
            display: 'block',
        },
        settingItem: {
            marginBottom: '24px',
        },
        settingLabel: {
            fontWeight: 600,
            marginBottom: '8px',
            display: 'block',
        },
        settingDescription: {
            color: theme.palette.neutralSecondary,
            fontSize: '13px',
            marginTop: '4px',
        },
        buttonGroup: {
            display: 'flex',
            gap: '12px',
            marginTop: '24px',
        },
        loadingContainer: {
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            height: '200px',
            color: theme.palette.neutralSecondary,
        },
    });

    return (
        <div className={classNames.container}>
            {/* Tab Header */}
            <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Notification Settings
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Configure notification preferences
                    </Text>
                </div>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={onRefresh}
                        title="Refresh data"
                        ariaLabel="Refresh data"
                    />
                </div>
            </div>

            {saveMessage && (
                <MessageBar
                    messageBarType={saveMessage.includes('success') ? MessageBarType.success : MessageBarType.error}
                    onDismiss={() => setSaveMessage(null)}
                    style={{ marginBottom: '16px' }}
                >
                    {saveMessage}
                </MessageBar>
            )}

            {loading ? (
                <div className={classNames.loadingContainer}>Loading notification settings...</div>
            ) : (
                <>
                    {/* Notification Channels */}
                    <div className={classNames.section}>
                        <Text className={classNames.sectionTitle}>Notification Channels</Text>

                        <div className={classNames.settingItem}>
                            <Toggle
                                label="Email Notifications"
                                checked={settings.EmailNotificationEnabled}
                                onChange={(_, checked) => setSettings(prev => ({ ...prev, EmailNotificationEnabled: checked || false }))}
                            />
                            <Text variant="small" className={classNames.settingDescription}>
                                Receive notifications via email for important updates and approvals
                            </Text>
                        </div>

                        <div className={classNames.settingItem}>
                            <Toggle
                                label="Microsoft Teams Notifications"
                                checked={settings.TeamsNotificationEnabled}
                                onChange={(_, checked) => setSettings(prev => ({ ...prev, TeamsNotificationEnabled: checked || false }))}
                            />
                            <Text variant="small" className={classNames.settingDescription}>
                                Receive notifications via Microsoft Teams
                            </Text>
                        </div>
                    </div>

                    {/* Approval Reminders */}
                    <div className={classNames.section}>
                        <Text className={classNames.sectionTitle}>Approval Reminders</Text>

                        <div className={classNames.settingItem}>
                            <Label>Reminder Frequency</Label>
                            <Slider
                                min={1}
                                max={14}
                                step={1}
                                value={settings.ApprovalReminderDays}
                                onChange={(value) => setSettings(prev => ({ ...prev, ApprovalReminderDays: value }))}
                                showValue
                                valueFormat={(value) => `${value} day${value > 1 ? 's' : ''}`}
                            />
                            <Text variant="small" className={classNames.settingDescription}>
                                How often to send reminders for pending approvals
                            </Text>
                        </div>
                    </div>

                    {/* Action Buttons */}
                    <div className={classNames.buttonGroup}>
                        <PrimaryButton
                            text="Save Settings"
                            iconProps={{ iconName: 'Save' }}
                            onClick={handleSave}
                            disabled={saving}
                        />
                        <DefaultButton
                            text="Reset to Defaults"
                            iconProps={{ iconName: 'Undo' }}
                            onClick={handleReset}
                            disabled={saving}
                        />
                    </div>
                </>
            )}
        </div>
    );
};

export default SettingsNotifications;
