import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    PrimaryButton,
    IconButton,
    TextField,
    IDropdownOption,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import EnhancedDataGrid from './EnhancedDataGrid';
import { IDataGridColumn } from './EnhancedDataGrid';
import MaterialForm from './MaterialForm';
import MaterialDetailsPanel from './MaterialDetailsPanel';
import { Tag, MonoText } from './TagRenderer';

export interface IMaterialsModuleProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    userDisplayName: string;
}

export interface IMaterial {
    ID: number;
    Material_Code: string;
    Material_Name: string;
    Category: string;
    UOM: string;
    Standard_Cost: number;
    Active: boolean;
    SubType?: string;
    Size?: string;
    MinStockLevel?: number;
    QRCodeURL?: string;
    qrcodeurl?: string;
    Specification?: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    [key: string]: any;
}

const MaterialsModule: React.FC<IMaterialsModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
}) => {
  const theme = getTheme();
  const [selectedMaterial, setSelectedMaterial] = React.useState<IMaterial | undefined>(undefined);
  const [isDetailsPanelOpen, setIsDetailsPanelOpen] = React.useState(false);
  const [isFormPanelOpen, setIsFormPanelOpen] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [panelKey, setPanelKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);

    React.useEffect(() => {
        const handleResize = (): void => {
            setIsMobileView(window.innerWidth < 768);
        };

        window.addEventListener('resize', handleResize);
        return (): void => {
            window.removeEventListener('resize', handleResize);
        };
    }, []);

    React.useEffect(() => {
        const handleGlobalRefresh = (): void => {
            setRefreshKey((prev) => prev + 1);
        };
        window.addEventListener('refreshData', handleGlobalRefresh);
        return (): void => {
            window.removeEventListener('refreshData', handleGlobalRefresh);
        };
    }, []);

    const classNames = mergeStyleSets({
        root: {
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            backgroundColor: theme.palette.white,
        },
        header: {
            marginBottom: '20px',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '12px',
        },
        headerTitle: {
            flex: 1,
            minWidth: '200px',
        },
        headerActions: {
            display: 'flex',
            gap: '8px',
            flexWrap: 'wrap',
        },

        gridContainer: {
            flex: 1,
            overflow: 'hidden',
            display: 'flex',
            flexDirection: 'column',
        },
        cardContainer: {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
            gap: '16px',
            overflow: 'auto',
            padding: '8px',
        },
        card: {
            padding: '16px',
            border: `1px solid ${theme.palette.neutralLight}`,
            borderRadius: '4px',
            backgroundColor: theme.palette.white,
            boxShadow: theme.effects.elevation4,
            cursor: 'pointer',
            transition: 'all 0.2s ease',
            selectors: {
                '&:hover': {
                    boxShadow: theme.effects.elevation8,
                    transform: 'translateY(-2px)',
                },
            },
        },
        cardField: {
            marginBottom: '12px',
            paddingBottom: '12px',
            borderBottom: `1px solid ${theme.palette.neutralLighter}`,
            selectors: {
                '&:last-child': {
                    borderBottom: 'none',
                    marginBottom: 0,
                    paddingBottom: 0,
                },
            },
        },
        cardLabel: {
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
            fontSize: '12px',
            marginBottom: '4px',
        },
        cardValue: {
            color: theme.palette.neutralSecondary,
            fontSize: '14px',
        },
        statusBadge: {
            display: 'inline-block',
            padding: '4px 8px',
            borderRadius: '3px',
            fontSize: '12px',
            fontWeight: 600,
        },
    });

    const materialColumns: IDataGridColumn[] = [
        {
            key: 'Material_Code',
            name: 'Material Code',
            fieldName: 'Material_Code',
            minWidth: 140,
            isResizable: true,
            onRender: (item: IMaterial) => <MonoText>{item.Material_Code}</MonoText>,
        },
        {
            key: 'Material_Name',
            name: 'Material Name',
            fieldName: 'Material_Name',
            minWidth: 180,
            isResizable: true,
        },
        {
            key: 'Category',
            name: 'Category',
            fieldName: 'Category',
            minWidth: 120,
            isResizable: true,
            onRender: (item: IMaterial) => <Tag text={item.Category} />,
        },
        {
            key: 'SubType',
            name: 'Sub-Type',
            fieldName: 'SubType',
            minWidth: 100,
            isResizable: true,
        },
        {
            key: 'Size',
            name: 'Size',
            fieldName: 'Size',
            minWidth: 80,
            isResizable: true,
        },
        {
            key: 'UOM',
            name: 'UOM',
            fieldName: 'UOM',
            minWidth: 80,
            isResizable: true,
        },
        {
            key: 'Standard_Cost',
            name: 'Std Cost',
            fieldName: 'Standard_Cost',
            minWidth: 100,
            isResizable: true,
        },
        {
            key: 'Active',
            name: 'Status',
            fieldName: 'Active',
            minWidth: 90,
            isResizable: true,
            onRender: (item: IMaterial) => <Tag text={item.Active ? 'Active' : 'Inactive'} />,
        },
    ];

    const handleRefresh = (): void => {
        setRefreshKey((prev) => prev + 1);
    };

    const handleNewMaterial = (): void => {
        setIsFormPanelOpen(true);
    };

  const handleRowSelected = (material: IMaterial): void => {
    setSelectedMaterial(material);
    setPanelKey((k) => k + 1);
    setIsDetailsPanelOpen(true);
  };

  const handleFormSubmit = (): void => {
    setIsFormPanelOpen(false);
    handleRefresh();
  };

  const closeDetailsPanel = (): void => {
    setSelectedMaterial(undefined);
    setIsDetailsPanelOpen(false);
  };

    return (
        <div className={classNames.root}>
            {/* Header Section */}
            <div className={classNames.header}>
                <div className={classNames.headerTitle}>
                    <Text variant="xxLarge" block style={{ fontWeight: 600, marginBottom: '4px' }}>
                        Materials Master
                    </Text>
                    <Text variant="medium" block style={{ color: theme.palette.neutralSecondary }}>
                        Manage and track all materials
                    </Text>
                </div>
                <div className={classNames.headerActions}>
                    <PrimaryButton
                        text="+ Add Material"
                        onClick={handleNewMaterial}
                        iconProps={{ iconName: 'Add' }}
                    />
                </div>
            </div>

            {/* Grid/Card Container */}
            <div className={classNames.gridContainer}>
          <EnhancedDataGrid
            key={`materials-${refreshKey}`}
            listName="ENT_Materials_Master"
            columns={materialColumns}
            pageSize={20}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            onRowSelected={handleRowSelected}
            showExport
          />
            </div>

            {/* Material Form Panel */}
            <MaterialForm
              isOpen={isFormPanelOpen}
              onDismiss={() => setIsFormPanelOpen(false)}
              onSubmitSuccess={handleFormSubmit}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
            />

            {/* Material Details Panel */}
            <MaterialDetailsPanel
              key={panelKey}
              isOpen={isDetailsPanelOpen}
              material={selectedMaterial}
              onDismiss={closeDetailsPanel}
              onRefresh={handleRefresh}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
            />
        </div>
      );
    };

export default MaterialsModule;
