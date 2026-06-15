/**
 * Role-Based Access Control (RBAC) Models
 * Defines the module/submodule structure and permission types
 */

// Module Keys - these correspond to the navigation keys in EnterpriseLayout
export type ModuleKey =
    | 'dashboard'
    | 'executive'
    | 'projects'
    | 'masterdata'
    | 'procurement'
    | 'finance'
    | 'reports'
    | 'settings';

// Submodule keys for Procurement
export type ProcurementSubModuleKey =
    | 'material-request'
    | 'purchase-requisition'
    | 'purchase-order'
    | 'goods-received-note';

// Submodule keys for Finance
export type FinanceSubModuleKey =
    | 'payment-requests'
    | 'approved-payments'
    | 'budget-tracking'
    | 'expense-register';

// Submodule keys for Master Data
export type MasterDataSubModuleKey =
    | 'materials'
    | 'vendors';

// Submodule keys for Reports
export type ReportsSubModuleKey =
    | 'procurement-reports'
    | 'finance-reports'
    | 'project-reports';

// Submodule keys for Settings
export type SettingsSubModuleKey =
    | 'approval-matrix'
    | 'system-config'
    | 'user-roles'
    | 'notifications';

export type SubModuleKey =
    | ProcurementSubModuleKey
    | FinanceSubModuleKey
    | MasterDataSubModuleKey
    | ReportsSubModuleKey
    | SettingsSubModuleKey;

// Permission levels
export type PermissionLevel = 'none' | 'view' | 'create' | 'edit' | 'delete' | 'full';

// Module definition
export interface IModuleDefinition {
    key: ModuleKey;
    name: string;
    icon: string;
    description: string;
    subModules?: ISubModuleDefinition[];
    requiredPermission: PermissionLevel;
}

// Submodule definition
export interface ISubModuleDefinition {
    key: SubModuleKey;
    parentModule: ModuleKey;
    name: string;
    description: string;
    requiredPermission: PermissionLevel;
}

// User role definition
export interface IUserRole {
    ID?: number;
    UserId?: number;
    User?: {
        ID: number;
        Title: string;
        Email: string;
    };
    Role: string;
    Department: string;
    Modules?: string; // JSON string of module permissions
    Permissions?: string; // Legacy field for description
}

// Role permission mapping
export interface IRolePermission {
    roleKey: string;
    roleName: string;
    modulePermissions: IModulePermission[];
}

// Module permission for a specific role
export interface IModulePermission {
    moduleKey: ModuleKey;
    subModuleKeys?: SubModuleKey[];
    permissionLevel: PermissionLevel;
    allowed?: boolean;
}

// User permissions context
export interface IUserPermissions {
    userId: number;
    userDisplayName: string;
    userEmail: string;
    role: string;
    department: string;
    modules: ModuleKey[];
    subModules: SubModuleKey[];
    permissions: Map<ModuleKey, PermissionLevel>;
    hasPermission: (moduleKey: ModuleKey, requiredLevel?: PermissionLevel) => boolean;
    hasSubModulePermission: (subModuleKey: SubModuleKey) => boolean;
}

// Default role-based permissions configuration
export const DEFAULT_ROLE_PERMISSIONS: IRolePermission[] = [
    {
        roleKey: 'Admin',
        roleName: 'Administrator',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'full' },
            { moduleKey: 'executive', permissionLevel: 'full' },
            { moduleKey: 'projects', permissionLevel: 'full' },
            { moduleKey: 'masterdata', permissionLevel: 'full' },
            { moduleKey: 'procurement', permissionLevel: 'full' },
            { moduleKey: 'finance', permissionLevel: 'full' },
            { moduleKey: 'reports', permissionLevel: 'full' },
            { moduleKey: 'settings', permissionLevel: 'full' },
        ],
    },
    {
        roleKey: 'Project Manager',
        roleName: 'Project Manager',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'view' },
            { moduleKey: 'executive', permissionLevel: 'view' },
            { moduleKey: 'projects', permissionLevel: 'full' },
            { moduleKey: 'masterdata', permissionLevel: 'view' },
            { moduleKey: 'procurement', permissionLevel: 'view' },
            { moduleKey: 'finance', permissionLevel: 'view' },
            { moduleKey: 'reports', permissionLevel: 'view' },
            { moduleKey: 'settings', permissionLevel: 'none' },
        ],
    },
    {
        roleKey: 'Finance Manager',
        roleName: 'Finance Manager',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'view' },
            { moduleKey: 'executive', permissionLevel: 'view' },
            { moduleKey: 'projects', permissionLevel: 'view' },
            { moduleKey: 'masterdata', permissionLevel: 'view' },
            { moduleKey: 'procurement', permissionLevel: 'view' },
            { moduleKey: 'finance', permissionLevel: 'full' },
            { moduleKey: 'reports', permissionLevel: 'view' },
            { moduleKey: 'settings', permissionLevel: 'none' },
        ],
    },
    {
        roleKey: 'Procurement Officer',
        roleName: 'Procurement Officer',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'view' },
            { moduleKey: 'executive', permissionLevel: 'none' },
            { moduleKey: 'projects', permissionLevel: 'view' },
            { moduleKey: 'masterdata', permissionLevel: 'view' },
            { moduleKey: 'procurement', permissionLevel: 'full' },
            { moduleKey: 'finance', permissionLevel: 'view' },
            { moduleKey: 'reports', permissionLevel: 'view' },
            { moduleKey: 'settings', permissionLevel: 'none' },
        ],
    },
    {
        roleKey: 'Approver',
        roleName: 'Approver',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'view' },
            { moduleKey: 'executive', permissionLevel: 'none' },
            { moduleKey: 'projects', permissionLevel: 'none' },
            { moduleKey: 'masterdata', permissionLevel: 'none' },
            { moduleKey: 'procurement', permissionLevel: 'view' },
            { moduleKey: 'finance', permissionLevel: 'view' },
            { moduleKey: 'reports', permissionLevel: 'none' },
            { moduleKey: 'settings', permissionLevel: 'none' },
        ],
    },
    {
        roleKey: 'Viewer',
        roleName: 'Viewer',
        modulePermissions: [
            { moduleKey: 'dashboard', permissionLevel: 'view' },
            { moduleKey: 'executive', permissionLevel: 'view' },
            { moduleKey: 'projects', permissionLevel: 'view' },
            { moduleKey: 'masterdata', permissionLevel: 'view' },
            { moduleKey: 'procurement', permissionLevel: 'view' },
            { moduleKey: 'finance', permissionLevel: 'view' },
            { moduleKey: 'reports', permissionLevel: 'view' },
            { moduleKey: 'settings', permissionLevel: 'none' },
        ],
    },
];

// All available modules in the system
export const MODULE_DEFINITIONS: IModuleDefinition[] = [
    {
        key: 'dashboard',
        name: 'Dashboard',
        icon: 'Home',
        description: 'Main dashboard with overview metrics',
        requiredPermission: 'view',
    },
    {
        key: 'executive',
        name: 'Executive Dashboard',
        icon: 'Chart',
        description: 'Executive-level analytics and reports',
        requiredPermission: 'view',
    },
    {
        key: 'projects',
        name: 'Projects',
        icon: 'ProjectCollection',
        description: 'Project management and tracking',
        requiredPermission: 'view',
    },
    {
        key: 'masterdata',
        name: 'Master Data',
        icon: 'Database',
        description: 'Manage materials, vendors, and other master data',
        requiredPermission: 'view',
        subModules: [
            { key: 'materials', parentModule: 'masterdata', name: 'Materials', description: 'Manage materials catalog', requiredPermission: 'edit' },
            { key: 'vendors', parentModule: 'masterdata', name: 'Vendors', description: 'Manage vendor information', requiredPermission: 'edit' },
        ],
    },
    {
        key: 'procurement',
        name: 'Procurement',
        icon: 'ShoppingCart',
        description: 'Procurement processes and workflows',
        requiredPermission: 'view',
        subModules: [
            { key: 'material-request', parentModule: 'procurement', name: 'Material Request', description: 'Request materials', requiredPermission: 'create' },
            { key: 'purchase-requisition', parentModule: 'procurement', name: 'Purchase Requisition', description: 'Create purchase requisitions', requiredPermission: 'create' },
            { key: 'purchase-order', parentModule: 'procurement', name: 'Purchase Order', description: 'Manage purchase orders', requiredPermission: 'edit' },
            { key: 'goods-received-note', parentModule: 'procurement', name: 'Goods Received Note', description: 'Record goods received', requiredPermission: 'create' },
        ],
    },
    {
        key: 'finance',
        name: 'Finance',
        icon: 'Money',
        description: 'Financial management and payments',
        requiredPermission: 'view',
        subModules: [
            { key: 'payment-requests', parentModule: 'finance', name: 'Payment Requests', description: 'Manage payment requests', requiredPermission: 'create' },
            { key: 'approved-payments', parentModule: 'finance', name: 'Approved Payments', description: 'View approved payments', requiredPermission: 'view' },
            { key: 'budget-tracking', parentModule: 'finance', name: 'Budget Tracking', description: 'Track budget utilization', requiredPermission: 'view' },
            { key: 'expense-register', parentModule: 'finance', name: 'Expense Register', description: 'Manage expense records', requiredPermission: 'create' },
        ],
    },
    {
        key: 'reports',
        name: 'Reports',
        icon: 'BarChartVertical',
        description: 'Generate and view reports',
        requiredPermission: 'view',
        subModules: [
            { key: 'procurement-reports', parentModule: 'reports', name: 'Procurement Reports', description: 'Procurement analytics', requiredPermission: 'view' },
            { key: 'finance-reports', parentModule: 'reports', name: 'Finance Reports', description: 'Financial analytics', requiredPermission: 'view' },
            { key: 'project-reports', parentModule: 'reports', name: 'Project Reports', description: 'Project analytics', requiredPermission: 'view' },
        ],
    },
    {
        key: 'settings',
        name: 'Settings',
        icon: 'Settings',
        description: 'System configuration and management',
        requiredPermission: 'edit',
        subModules: [
            { key: 'approval-matrix', parentModule: 'settings', name: 'Approval Matrix', description: 'Configure approval workflows', requiredPermission: 'full' },
            { key: 'system-config', parentModule: 'settings', name: 'System Configuration', description: 'System settings', requiredPermission: 'full' },
            { key: 'user-roles', parentModule: 'settings', name: 'User Roles', description: 'Manage user roles and permissions', requiredPermission: 'full' },
            { key: 'notifications', parentModule: 'settings', name: 'Notifications', description: 'Configure notifications', requiredPermission: 'full' },
        ],
    },
];

// Helper function to check if a permission level meets the required level
export function hasPermissionLevel(userLevel: PermissionLevel, requiredLevel: PermissionLevel): boolean {
    const levelHierarchy: PermissionLevel[] = ['none', 'view', 'create', 'edit', 'delete', 'full'];
    const userIndex = levelHierarchy.indexOf(userLevel);
    const requiredIndex = levelHierarchy.indexOf(requiredLevel);
    return userIndex >= requiredIndex;
}

// Helper to get module definition by key
export function getModuleDefinition(key: ModuleKey): IModuleDefinition | undefined {
    return MODULE_DEFINITIONS.find(m => m.key === key);
}

// Helper to get submodule definition by key
export function getSubModuleDefinition(key: SubModuleKey): ISubModuleDefinition | undefined {
    for (const module of MODULE_DEFINITIONS) {
        if (module.subModules) {
            const subModule = module.subModules.find(s => s.key === key);
            if (subModule) return subModule;
        }
    }
    return undefined;
}
