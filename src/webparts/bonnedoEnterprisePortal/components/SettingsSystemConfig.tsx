import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    Toggle,
    TextField,
    PrimaryButton,
    IconButton,
    DefaultButton,
    MessageBar,
    MessageBarType,
    Label,
    Dropdown,
    IDropdownOption,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../services/SharePointService';

// Settings System Config Component
interface ISettingsSystemConfigProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    onRefresh?: () => void;
    isMobileView?: boolean;
}

const SettingsSystemConfig: React.FC<ISettingsSystemConfigProps> = ({
    spHttpClient,
    pageContext,
    onRefresh,
    isMobileView = false
}) => {
    const theme = getTheme();
    const [settings, setSettings] = React.useState({
        ApplicationName: 'Bonnedo Enterprise Portal',
        ApplicationTheme: 'blue',
        DefaultCurrency: 'NGN',
        ProcurementApprovalLimit: '500000',
        AutoCreatePR: false,
        AutoCreatePO: false,
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
                const data = await sharePointService.getListData('SYS_System_Config', undefined, 10);
                if (data.length > 0) {
                    const config = data[0];
                    setSettings({
                        ApplicationName: config.ApplicationName || 'Bonnedo Enterprise Portal',
                        ApplicationTheme: config.ApplicationTheme || 'blue',
                        DefaultCurrency: config.DefaultCurrency || 'NGN',
                        ProcurementApprovalLimit: config.ProcurementApprovalLimit?.toString() || '500000',
                        AutoCreatePR: config.AutoCreatePR || false,
                        AutoCreatePO: config.AutoCreatePO || false,
                    });
                }
            } catch (error) {
                console.error('Error fetching system config:', error);
            }
            setLoading(false);
        };
        fetchSettings();
    }, [sharePointService]);

    const handleSave = async (): Promise<void> => {
        setSaving(true);
        try {
            const webUrl = pageContext.web.absoluteUrl;
            // Get existing config first
            const existing = await sharePointService.getListData('SYS_System_Config', undefined, 1);

            if (existing.length > 0) {
                // Update existing
                const url = `${webUrl}/_api/web/lists/getByTitle('SYS_System_Config')/items(${existing[0].ID})`;
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
                // Create new
                await sharePointService.createListItem('SYS_System_Config', settings);
            }

            setSaveMessage('Settings saved successfully!');
            setTimeout(() => setSaveMessage(null), 3000);
        } catch (error) {
            setSaveMessage('Failed to save settings');
        }
        setSaving(false);
    };

    const handleReset = (): void => {
        setSettings({
            ApplicationName: 'Bonnedo Enterprise Portal',
            ApplicationTheme: 'blue',
            DefaultCurrency: 'NGN',
            ProcurementApprovalLimit: '500000',
            AutoCreatePR: false,
            AutoCreatePO: false,
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
            marginBottom: '20px',
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

    const themeOptions: IDropdownOption[] = [
        { key: 'blue', text: 'Blue (Microsoft Standard)' },
        { key: 'green', text: 'Green' },
        { key: 'purple', text: 'Purple' },
        { key: 'teal', text: 'Teal' },
    ];

    const currencyOptions: IDropdownOption[] = [
        { key: 'NGN', text: 'Nigerian Naira (₦)' },
        { key: 'USD', text: 'US Dollar ($)' },
        { key: 'EUR', text: 'Euro (€)' },
        { key: 'GBP', text: 'British Pound (£)' },
    ];

    return (
        <div className={classNames.container}>
            {/* Tab Header */}
            <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        System Configuration
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage system preferences and settings
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
                <div className={classNames.loadingContainer}>Loading settings...</div>
            ) : (
                <>
                    {/* Application Settings */}
                    <div className={classNames.section}>
                        <Text className={classNames.sectionTitle}>Application Settings</Text>

                        <div className={classNames.settingItem}>
                            <Label>Application Name</Label>
                            <TextField
                                value={settings.ApplicationName}
                                onChange={(_, value) => setSettings(prev => ({ ...prev, ApplicationName: value || '' }))}
                            />
                        </div>

                        <div className={classNames.settingItem}>
                            <Label>Application Theme</Label>
                            <Dropdown
                                options={themeOptions}
                                selectedKey={settings.ApplicationTheme}
                                onChange={(_, option) => setSettings(prev => ({ ...prev, ApplicationTheme: option?.key as string || 'blue' }))}
                            />
                        </div>
                    </div>

                    {/* Financial Settings */}
                    <div className={classNames.section}>
                        <Text className={classNames.sectionTitle}>Financial Settings</Text>

                        <div className={classNames.settingItem}>
                            <Label>Default Currency</Label>
                            <Dropdown
                                options={currencyOptions}
                                selectedKey={settings.DefaultCurrency}
                                onChange={(_, option) => setSettings(prev => ({ ...prev, DefaultCurrency: option?.key as string || 'NGN' }))}
                            />
                        </div>

                        <div className={classNames.settingItem}>
                            <Label>Procurement Approval Limit</Label>
                            <TextField
                                type="number"
                                prefix="₦"
                                value={settings.ProcurementApprovalLimit}
                                onChange={(_, value) => setSettings(prev => ({ ...prev, ProcurementApprovalLimit: value || '0' }))}
                            />
                        </div>
                    </div>

                    {/* Automation Settings */}
                    <div className={classNames.section}>
                        <Text className={classNames.sectionTitle}>Automation Settings</Text>

                        <div className={classNames.settingItem}>
                            <Toggle
                                label="Auto Create Purchase Requisition"
                                checked={settings.AutoCreatePR}
                                onChange={(_, checked) => setSettings(prev => ({ ...prev, AutoCreatePR: checked || false }))}
                            />
                            <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                                Automatically create PR when Material Request is approved
                            </Text>
                        </div>

                        <div className={classNames.settingItem}>
                            <Toggle
                                label="Auto Create Purchase Order"
                                checked={settings.AutoCreatePO}
                                onChange={(_, checked) => setSettings(prev => ({ ...prev, AutoCreatePO: checked || false }))}
                            />
                            <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                                Automatically create PO when Purchase Requisition is approved
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

export default SettingsSystemConfig;
