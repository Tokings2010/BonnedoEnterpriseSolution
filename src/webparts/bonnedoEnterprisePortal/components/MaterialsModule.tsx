import * as React from 'react';
import {
    Text,
    getTheme,
    mergeStyleSets,
    PrimaryButton,
    IconButton,
    TextField,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';
import MaterialForm from './MaterialForm';
import MaterialDetailsPanel from './MaterialDetailsPanel';

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
            minWidth: 120,
            isResizable: true,
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
            minWidth: 150,
            isResizable: true,
        },
        {
            key: 'UOM',
            name: 'Unit of Measure',
            fieldName: 'UOM',
            minWidth: 100,
            isResizable: true,
        },
        {
            key: 'Standard_Cost',
            name: 'Standard Cost',
            fieldName: 'Standard_Cost',
            minWidth: 120,
            isResizable: true,
        },
        {
            key: 'Active',
            name: 'Active',
            fieldName: 'Active',
            minWidth: 80,
            isResizable: true,
            onRender: (item: IMaterial) => {
                return (
                    <span style={{
                        color: item.Active ? theme.palette.green : theme.palette.red
                    }}>
                        {item.Active ? 'Yes' : 'No'}
                    </span>
                );
            }
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
    setIsDetailsPanelOpen(true);
  };

  const handleFormSubmit = (): void => {
    setIsFormPanelOpen(false);
    handleRefresh();
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
        {isMobileView ? (
          <DataGrid
            key={`materials-mobile-${refreshKey}`}
            listName="ENT_Materials_Master"
            columns={materialColumns}
            pageSize={20}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            onRowSelected={handleRowSelected}
          />
        ) : (
          <DataGrid
            key={`materials-desktop-${refreshKey}`}
            listName="ENT_Materials_Master"
            columns={materialColumns}
            pageSize={20}
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            onRowSelected={handleRowSelected}
          />
        )}
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
              isOpen={isDetailsPanelOpen}
              material={selectedMaterial}
              onDismiss={() => setIsDetailsPanelOpen(false)}
              onRefresh={handleRefresh}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
            />
        </div>
      );
    };

export default MaterialsModule;
