import * as React from 'react';
import {
  Nav,
  INavLinkGroup,
  Text,
  getTheme,
  IconButton,
  Breadcrumb,
  IBreadcrumbItem,
  INavLink,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ProcurementModule from './ProcurementModule';
import FinanceModule from './FinanceModule';
import DashboardModule from './DashboardModule';
import DashboardErrorBoundary from './DashboardErrorBoundary';
import ExecutiveDashboard from './dashboard/ExecutiveDashboard';
import ProjectsModule from './ProjectsModule';
import ReportsModule from './ReportsModule';
import SettingsModule from './SettingsModule';
import MaterialsManagementModule from './MaterialsManagementModule/MaterialsManagementModule';
import styles from './EnterpriseLayout.module.scss';
import { IUserPermissions, ModuleKey } from '../models/PermissionModels';

interface IEnterpriseLayoutProps {
  userDisplayName: string;
  isDarkTheme: boolean;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext: WebPartContext;
  portalMode?: boolean;
  fullWidth?: boolean;
  hasTeamsContext?: boolean;
  userPermissions?: IUserPermissions;
}

type TopMenuKey = 'dashboard' | 'executive' | 'projects' | 'material' | 'procurement' | 'finance' | 'reports' | 'settings';

// Access Denied component
const AccessDeniedMessage: React.FC = () => {
  const theme = getTheme();
  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      height: '100%',
      padding: '40px',
      textAlign: 'center'
    }}>
      <Text
        variant="xxLarge"
        style={{
          color: theme.palette.red,
          marginBottom: '16px'
        }}
      >
        Access Denied
      </Text>
      <Text variant="large" style={{ color: theme.palette.neutralSecondary }}>
        You don't have permission to access this module.
      </Text>
      <Text variant="medium" style={{ color: theme.palette.neutralTertiary, marginTop: '8px' }}>
        Please contact your administrator if you believe this is an error.
      </Text>
    </div>
  );
};

const EnterpriseLayout: React.FC<IEnterpriseLayoutProps> = ({
  userDisplayName,
  isDarkTheme,
  spHttpClient,
  pageContext,
  webPartContext,
  portalMode = false,
  fullWidth = false,
  hasTeamsContext = false,
  userPermissions
}) => {
  const [selectedTopMenu, setSelectedTopMenu] = React.useState<TopMenuKey>('dashboard');
  const [isNavCollapsed, setIsNavCollapsed] = React.useState(false);
  const [isRefreshing, setIsRefreshing] = React.useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = React.useState(false);
  const [isMobile, setIsMobile] = React.useState(false);

  // Handle window resize for responsive behavior
  React.useEffect(() => {
    const handleResize = (): void => {
      const mobile = window.innerWidth < 768;
      setIsMobile(mobile);
      if (mobile) {
        setIsNavCollapsed(true);
      }
    };

    // Set initial value
    handleResize();

    window.addEventListener('resize', handleResize);
    return () => {
      window.removeEventListener('resize', handleResize);
    };
  }, []);

  // Redirect to dashboard if no access to current module
  React.useEffect(() => {
    if (userPermissions && selectedTopMenu !== 'dashboard' && !userPermissions.modules.includes(selectedTopMenu as ModuleKey)) {
      setSelectedTopMenu('dashboard');
    }
  }, [selectedTopMenu, userPermissions]);

  // Effect to hide SharePoint/Teams chrome elements when portalMode is enabled
  React.useEffect(() => {
    if (portalMode || hasTeamsContext) {
      // Array of SharePoint chrome element selectors to hide
      const sharePointChromeSelectors = [
        '#spLeftNav',
        '#SuiteNavWrapper',
        '.od-SuiteNav',
        '#spCommandBar',
        '#sp-appBar',
        '.sp-appBar',
        '#O365_HeaderRow',
        '#o365cnb',
        '.o365cs-nav-header',
        '#nav_partialPlaceHolder',
        '[data-automation-id="suiteNavBar"]',
        '.sp-suiteNavBar'
      ];

      // Teams-specific selectors to hide
      const teamsChromeSelectors = [
        '.teams-horizontal-left-nav',
        '[data-automation-id="teamsLeftNav"]',
        '.tlight-commands',
        '#appbase-fluent-commands-layer'
      ];

      const allSelectors = portalMode ? [...sharePointChromeSelectors, ...teamsChromeSelectors] : teamsChromeSelectors;

      const hideChrome = (): void => {
        allSelectors.forEach(selector => {
          const elements = document.querySelectorAll(selector);
          elements.forEach(element => {
            (element as HTMLElement).style.display = 'none';
            (element as HTMLElement).style.visibility = 'hidden';
          });
        });

        // Also set document overflow to hidden to prevent scrolling SharePoint/Teams
        document.body.style.overflow = 'hidden';
        document.documentElement.style.overflow = 'hidden';
      };

      // Apply immediately
      hideChrome();

      // Apply again after a short delay to catch dynamically loaded elements
      const timeoutId = setTimeout(hideChrome, 100);
      const timeoutId2 = setTimeout(hideChrome, 500);
      const timeoutId3 = setTimeout(hideChrome, 1000);

      return () => {
        clearTimeout(timeoutId);
        clearTimeout(timeoutId2);
        clearTimeout(timeoutId3);
        // Restore on unmount
        document.body.style.overflow = '';
        document.documentElement.style.overflow = '';
      };
    }
  }, [portalMode, hasTeamsContext]);

  // Refresh data function
  const refreshData = (): void => {
    setIsRefreshing(true);
    setTimeout(() => {
      setIsRefreshing(false);
      // Trigger refresh in child modules
      window.dispatchEvent(new CustomEvent('refreshData'));
    }, 1000);
  };

  // Define side navigation items - filtered based on user permissions
  const getSideNavItems = (): INavLinkGroup[] => {
    // Default navigation items
    const allItems = [
      { name: 'Dashboard', url: '#', icon: 'Home', key: 'dashboard' },
      { name: 'Executive Dashboard', url: '#', icon: 'Chart', key: 'executive' },
      { name: 'Projects', url: '#', icon: 'ProjectCollection', key: 'projects' },
      { name: 'Material', url: '#', icon: 'Package', key: 'material' },
      { name: 'Procurement', url: '#', icon: 'ShoppingCart', key: 'procurement' },
      { name: 'Finance', url: '#', icon: 'Money', key: 'finance' },
      { name: 'Reports', url: '#', icon: 'BarChartVertical', key: 'reports' },
      { name: 'Settings', url: '#', icon: 'Settings', key: 'settings' },
    ];

    // Filter based on user permissions
    if (userPermissions && userPermissions.modules) {
      const allowedModules = userPermissions.modules;
      const filteredItems = allItems.filter(item =>
        allowedModules.includes(item.key as ModuleKey)
      );
      return [{ links: filteredItems }];
    }

    // If no permissions, show all items (backward compatibility)
    return [{ links: allItems }];
  };

  // Get breadcrumb items based on current module
  const getBreadcrumbItems = (): IBreadcrumbItem[] => {
    const breadcrumbs: Record<TopMenuKey, IBreadcrumbItem[]> = {
      dashboard: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Dashboard', key: 'dashboard', isCurrentItem: true },
      ],
      executive: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Executive Dashboard', key: 'executive', isCurrentItem: true },
      ],
      projects: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Projects', key: 'projects', isCurrentItem: true },
      ],
      material: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Material', key: 'material', isCurrentItem: true },
      ],
      procurement: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Procurement', key: 'procurement', isCurrentItem: true },
      ],
      finance: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Finance', key: 'finance', isCurrentItem: true },
      ],
      reports: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Reports', key: 'reports', isCurrentItem: true },
      ],
      settings: [
        { text: 'Home', key: 'home', onClick: () => setSelectedTopMenu('dashboard') },
        { text: 'Settings', key: 'settings', isCurrentItem: true },
      ],
    };

    return breadcrumbs[selectedTopMenu] || [];
  };

  // Render module based on selected menu
  const renderModule = (): React.ReactElement => {
    // Check if user has permission to access the module
    const hasModuleAccess = (moduleKey: string): boolean => {
      if (!userPermissions) return true; // Backward compatibility
      return userPermissions.modules.includes(moduleKey as ModuleKey);
    };

    // If user doesn't have access, redirect to dashboard
    if (!hasModuleAccess(selectedTopMenu) && selectedTopMenu !== 'dashboard') {
      setSelectedTopMenu('dashboard');
      return <div>Redirecting...</div>;
    }

    switch (selectedTopMenu) {
      case 'dashboard':
        return (
          <DashboardErrorBoundary>
            <DashboardModule
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              userDisplayName={userDisplayName}
              onNavigate={(module: string) => setSelectedTopMenu(module as TopMenuKey)}
            />
          </DashboardErrorBoundary>
        );
      case 'executive':
        if (!hasModuleAccess('executive')) {
          return <AccessDeniedMessage />;
        }
        return (
          <ExecutiveDashboard
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
          />
        );
      case 'projects':
        if (!hasModuleAccess('projects')) {
          return <AccessDeniedMessage />;
        }
        return (
          <ProjectsModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            webPartContext={webPartContext}
          />
        );
      case 'material':
        if (!hasModuleAccess('material')) {
          return <AccessDeniedMessage />;
        }
        return (
          <MaterialsManagementModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            userPermissions={userPermissions}
          />
        );
      case 'procurement':
        if (!hasModuleAccess('procurement')) {
          return <AccessDeniedMessage />;
        }
        return (
          <ProcurementModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            webPartContext={webPartContext}
          />
        );
      case 'finance':
        if (!hasModuleAccess('finance')) {
          return <AccessDeniedMessage />;
        }
        return (
          <FinanceModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            webPartContext={webPartContext}
          />
        );
      case 'reports':
        if (!hasModuleAccess('reports')) {
          return <AccessDeniedMessage />;
        }
        return (
          <ReportsModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
          />
        );
      case 'settings':
        // Only Admin users can access Settings
        if (!userPermissions || userPermissions.role !== 'Admin') {
          return <AccessDeniedMessage />;
        }
        return (
          <SettingsModule
            spHttpClient={spHttpClient}
            pageContext={pageContext}
            userDisplayName={userDisplayName}
            webPartContext={webPartContext}
          />
        );
      default:
        return <div>Module not found</div>;
    }
  };

  // Handle navigation toggle
  const toggleNav = (): void => {
    if (isMobile) {
      setIsMobileMenuOpen(!isMobileMenuOpen);
    } else {
      setIsNavCollapsed(!isNavCollapsed);
    }
  };

  // Determine root class based on fullWidth or portalMode
  const rootClass = fullWidth || portalMode ? `${styles.root} ${styles.rootFullWidth}` : styles.root;

  return (
    <div className={rootClass}>
      {/* Header */}
      <header className={styles.header}>
        <div className={styles.headerContent}>
          <div className={styles.logoSection}>
            <IconButton
              iconProps={{ iconName: 'GlobalNavButton' }}
              onClick={toggleNav}
              className={styles.iconButton}
              ariaLabel="Toggle navigation"
            />
            <div className={styles.logo}>B</div>
            <Text className={styles.portalTitle}>Bonnedo Enterprise Portal</Text>
          </div>
          <div className={styles.headerActions}>
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={refreshData}
              disabled={isRefreshing}
              className={styles.iconButton}
              ariaLabel="Refresh data"
              title="Refresh"
            />
            <IconButton
              iconProps={{ iconName: 'Contact' }}
              text={userDisplayName}
              className={styles.iconButton}
              ariaLabel="User profile"
            />
          </div>
        </div>
      </header>

      {/* Main Container */}
      <div className={styles.mainContainer}>
        {/* Side Navigation */}
        <nav
          className={`${styles.sideNav} ${isNavCollapsed ? styles.collapsed : ''} ${isMobileMenuOpen ? styles.mobileOpen : ''}`}
          style={{ width: isNavCollapsed ? 70 : 260 }}
        >
          <Nav
            groups={getSideNavItems()}
            selectedKey={selectedTopMenu}
            ariaLabel="Side Navigation"
            onLinkClick={(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
              if (item?.key) {
                setSelectedTopMenu(item.key as TopMenuKey);
                if (isMobile) setIsMobileMenuOpen(false);
              }
            }}
            styles={{
              root: {
                width: '100%',
              },
              link: {
                padding: isNavCollapsed ? '8px 4px' : '8px 16px',
                fontSize: isNavCollapsed ? '12px' : '14px',
              },
            }}
          />
        </nav>

        {/* Main Content Area */}
        <main className={styles.mainContent}>
          {/* Breadcrumb Header */}
          <div className={styles.contentHeader}>
            <Breadcrumb items={getBreadcrumbItems()} />
          </div>

          {/* Content Area */}
          <div className={styles.contentArea}>
            {renderModule()}
          </div>
        </main>
      </div>
    </div>
  );
};

export default EnterpriseLayout;
