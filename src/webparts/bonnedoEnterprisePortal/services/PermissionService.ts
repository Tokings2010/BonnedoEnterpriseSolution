import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
    IUserPermissions,
    IUserRole,
    ModuleKey,
    SubModuleKey,
    PermissionLevel,
    DEFAULT_ROLE_PERMISSIONS,
    MODULE_DEFINITIONS,
    hasPermissionLevel,
    IRolePermission,
    IModulePermission
} from '../models/PermissionModels';
import { SharePointService, IListItem } from './SharePointService';

/**
 * Permission Service - handles role-based access control
 * Fetches user permissions from SharePoint and provides authorization checks
 */
export class PermissionService {
    private spHttpClient: SPHttpClient;
    private pageContext: PageContext;
    private sharePointService: SharePointService;
    private cachedPermissions: IUserPermissions | null = null;
    private currentUserId: number | null = null;

    constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
        this.spHttpClient = spHttpClient;
        this.pageContext = pageContext;
        this.sharePointService = new SharePointService(spHttpClient, pageContext);
    }

    /**
     * Get the current user's permissions from SharePoint
     * Falls back to Admin permissions if no role is assigned (backward compatibility)
     * Site owners automatically get Admin access
     */
    public async getUserPermissions(): Promise<IUserPermissions> {
        // Return cached permissions if available
        if (this.cachedPermissions) {
            return this.cachedPermissions;
        }

        const currentUserId = this.pageContext.legacyPageContext.userId;
        const userDisplayName = this.pageContext.user.displayName;
        const userEmail = this.pageContext.user.email;

        try {
            // First, check if user is a site owner - they get Admin access automatically
            // Using SPFx built-in permission check which is more reliable
            const isSiteOwner = await this.checkIfSiteOwner();
            if (isSiteOwner) {
                console.log('User is a site owner - granting Admin access');

                // Auto-save site owner to SYS_User_Roles list
                await this.autoSaveSiteOwnerToRoles(currentUserId, userDisplayName, userEmail);

                this.cachedPermissions = this.buildDefaultPermissions(
                    currentUserId,
                    userDisplayName,
                    userEmail,
                    'Admin'
                );
                return this.cachedPermissions;
            }

            // Try to find user role in SharePoint
            const userRoles = await this.sharePointService.getUserRoles();

            // Find the role for the current user
            const userRole = userRoles.find((role: IListItem) =>
                role.UserId === currentUserId || role.User?.ID === currentUserId
            );

            if (userRole) {
                // User has an assigned role
                this.cachedPermissions = this.buildUserPermissions(
                    currentUserId,
                    userDisplayName,
                    userEmail,
                    userRole as unknown as IUserRole
                );
            } else {
                // No role assigned - use Viewer as default
                this.cachedPermissions = this.buildDefaultPermissions(
                    currentUserId,
                    userDisplayName,
                    userEmail,
                    'Viewer'
                );
            }

            return this.cachedPermissions;
        } catch (error) {
            console.error('Error fetching user permissions:', error);
            // Return default Viewer permissions on error
            return this.buildDefaultPermissions(
                currentUserId,
                userDisplayName,
                userEmail,
                'Viewer'
            );
        }
    }

    /**
     * Check if current user is a site owner using SPFx built-in isSiteAdmin property
     * This is the most reliable and simple approach
     */
    private async checkIfSiteOwner(): Promise<boolean> {
        try {
            // Method 1: Use SPFx built-in isSiteAdmin property (most reliable and simple!)
            // This directly tells us if the user is a site administrator/owner
            const isSiteAdmin = this.pageContext.legacyPageContext.isSiteAdmin;
            console.log('SPFx isSiteAdmin:', isSiteAdmin);

            if (isSiteAdmin) {
                console.log('User is a site admin - granting Admin access');
                return true;
            }

            // Method 2: Check if user is in Owners group via REST API
            const webUrl = this.pageContext.web.absoluteUrl;
            const currentUserId = this.pageContext.legacyPageContext.userId;

            try {
                // Get all site groups and check if user is in Owners
                const groupsResponse = await this.spHttpClient.get(
                    `${webUrl}/_api/web/sitegroups`,
                    SPHttpClient.configurations.v1
                );

                if (groupsResponse.ok) {
                    const groupsData = await groupsResponse.json();
                    const groups = groupsData.value || [];

                    // Find the Owners group
                    const ownersGroup = groups.find((g: { Title: string }) =>
                        g.Title && g.Title.toLowerCase() === 'owners'
                    );

                    if (ownersGroup) {
                        console.log('Found Owners group, ID:', ownersGroup.Id);

                        // Check if current user is in this group
                        const checkUserUrl = `${webUrl}/_api/web/sitegroups(${ownersGroup.Id})/users(${currentUserId})`;
                        const userInGroupResponse = await this.spHttpClient.get(
                            checkUserUrl,
                            SPHttpClient.configurations.v1
                        );

                        if (userInGroupResponse.ok) {
                            console.log('User is in Owners group - granting Admin access');
                            return true;
                        }
                    }
                }
            } catch (groupError) {
                console.warn('Error checking group membership via REST:', groupError);
            }

            console.log('User is not a site owner');
            return false;
        } catch (error) {
            console.error('Error checking site ownership:', error);
            return false;
        }
    }

    /**
     * Auto-save site owner to SYS_User_Roles list if not already present
     * This ensures site owners are properly recorded in the system
     */
    private async autoSaveSiteOwnerToRoles(
        currentUserId: number,
        userDisplayName: string,
        userEmail: string
    ): Promise<void> {
        try {
            // Save user as Admin in SYS_User_Roles
            await this.sharePointService.saveUserRole(
                currentUserId,
                userDisplayName,
                'Admin',
                '',
                'Auto-assigned: Site Owner'
            );
            console.log('Site owner auto-saved to SYS_User_Roles with Admin role');
        } catch (error) {
            // Non-fatal error - log but don't break the flow
            console.warn('Failed to auto-save site owner to SYS_User_Roles:', error);
        }
    }

    // ========== SharePoint Group Management Methods ==========

    /**
     * Add a user to the SharePoint Owners group
     * @param userId - The SharePoint user ID
     */
    public async addUserToOwnersGroup(userId: number): Promise<boolean> {
        try {
            const webUrl = this.pageContext.web.absoluteUrl;
            // Get the Owners group
            const groupsResponse = await this.spHttpClient.get(
                `${webUrl}/_api/web/sitegroups?$filter=Title%20eq%20'Owners'`,
                SPHttpClient.configurations.v1
            );

            if (!groupsResponse.ok) {
                console.error('Failed to get Owners group');
                return false;
            }

            const groupsData = await groupsResponse.json();
            if (!groupsData.value || groupsData.value.length === 0) {
                console.error('Owners group not found');
                return false;
            }

            const ownersGroupId = groupsData.value[0].Id;

            // Add user to the group
            const addResponse = await this.spHttpClient.post(
                `${webUrl}/_api/web/sitegroups(${ownersGroupId})/users`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        __metadata: { type: 'SP.User' },
                        LoginName: `i:0#.f|membership|${userId}`,
                    }),
                }
            );

            if (!addResponse.ok) {
                const errorText = await addResponse.text();
                console.error('Failed to add user to Owners group:', errorText);
                return false;
            }

            console.log('User added to Owners group successfully');
            return true;
        } catch (error) {
            console.error('Error adding user to Owners group:', error);
            return false;
        }
    }

    /**
     * Add a user to the SharePoint Members group
     * @param userId - The SharePoint user ID
     */
    public async addUserToMembersGroup(userId: number): Promise<boolean> {
        try {
            const webUrl = this.pageContext.web.absoluteUrl;
            // Get the Members group
            const groupsResponse = await this.spHttpClient.get(
                `${webUrl}/_api/web/sitegroups?$filter=Title%20eq%20'Members'`,
                SPHttpClient.configurations.v1
            );

            if (!groupsResponse.ok) {
                console.error('Failed to get Members group');
                return false;
            }

            const groupsData = await groupsResponse.json();
            if (!groupsData.value || groupsData.value.length === 0) {
                console.error('Members group not found');
                return false;
            }

            const membersGroupId = groupsData.value[0].Id;

            // Add user to the group
            const addResponse = await this.spHttpClient.post(
                `${webUrl}/_api/web/sitegroups(${membersGroupId})/users`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        __metadata: { type: 'SP.User' },
                        LoginName: `i:0#.f|membership|${userId}`,
                    }),
                }
            );

            if (!addResponse.ok) {
                const errorText = await addResponse.text();
                console.error('Failed to add user to Members group:', errorText);
                return false;
            }

            console.log('User added to Members group successfully');
            return true;
        } catch (error) {
            console.error('Error adding user to Members group:', error);
            return false;
        }
    }

    /**
     * Remove a user from SharePoint site groups
     * @param userId - The SharePoint user ID
     */
    public async removeUserFromSite(userId: number): Promise<boolean> {
        try {
            const webUrl = this.pageContext.web.absoluteUrl;

            // Get all site groups
            const groupsResponse = await this.spHttpClient.get(
                `${webUrl}/_api/web/sitegroups`,
                SPHttpClient.configurations.v1
            );

            if (!groupsResponse.ok) return false;

            const groupsData = await groupsResponse.json();
            const groups = groupsData.value || [];

            // Remove user from each group (Owners, Members, Visitors)
            for (const group of groups) {
                const groupTitle = group.Title?.toLowerCase() || '';
                if (groupTitle.includes('owner') ||
                    groupTitle.includes('member') ||
                    groupTitle.includes('visitor')) {

                    try {
                        await this.spHttpClient.post(
                            `${webUrl}/_api/web/sitegroups(${group.Id})/users(${userId})`,
                            SPHttpClient.configurations.v1,
                            {
                                headers: {
                                    'Accept': 'application/json',
                                    'X-HTTP-Method': 'DELETE',
                                    'If-Match': '*',
                                },
                            }
                        );
                    } catch (e) {
                        // User might not be in this group, continue
                        console.log(`User might not be in group ${group.Title}`);
                    }
                }
            }

            console.log('User removed from site groups successfully');
            return true;
        } catch (error) {
            console.error('Error removing user from site:', error);
            return false;
        }
    }

    /**
     * Assign a role to a user and update SharePoint group membership
     * @param userId - The SharePoint user ID
     * @param role - The role to assign
     * @param sharePointService - SharePoint service for database operations
     */
    public async assignRoleWithGroupMembership(
        userId: number,
        role: string,
        sharePointService: SharePointService
    ): Promise<boolean> {
        try {
            // First, check if user already has a role
            const existingRoles = await sharePointService.getUserRoles();
            const existingRole = existingRoles.find((r: IListItem) =>
                r.UserId === userId || r.User?.ID === userId
            );

            // If user already has a role and it's changing, remove from old group first
            if (existingRole && existingRole.Role !== role) {
                await this.removeUserFromSite(userId);
            }

            // Add user to appropriate SharePoint group based on role
            if (role === 'Admin' || role === 'Administrator') {
                // Add to Owners group
                await this.addUserToOwnersGroup(userId);
            } else {
                // Add to Members group for all other roles
                await this.addUserToMembersGroup(userId);
            }

            return true;
        } catch (error) {
            console.error('Error assigning role with group membership:', error);
            return false;
        }
    }

    /**
     * Revoke a user's role and remove from SharePoint site
     * @param userId - The SharePoint user ID
     * @param sharePointService - SharePoint service for database operations
     */
    public async revokeRoleAndGroupMembership(
        userId: number,
        sharePointService: SharePointService
    ): Promise<boolean> {
        try {
            // Remove from SharePoint groups
            await this.removeUserFromSite(userId);

            return true;
        } catch (error) {
            console.error('Error revoking role and group membership:', error);
            return false;
        }
    }

    /**
     * Build user permissions from SharePoint role record
     */
    private buildUserPermissions(
        userId: number,
        userDisplayName: string,
        userEmail: string,
        userRole: IUserRole
    ): IUserPermissions {
        const roleKey = userRole.Role;
        const department = userRole.Department || '';

        // Get role-based default permissions
        const rolePermission = DEFAULT_ROLE_PERMISSIONS.find(rp => rp.roleKey === roleKey);

        let modulePermissions: IModulePermission[];
        const allowedModules: ModuleKey[] = [];
        const allowedSubModules: SubModuleKey[] = [];

        if (rolePermission) {
            modulePermissions = rolePermission.modulePermissions;
        } else {
            // Fallback to Viewer permissions if role not found
            const viewerRole = DEFAULT_ROLE_PERMISSIONS.find(rp => rp.roleKey === 'Viewer');
            modulePermissions = viewerRole?.modulePermissions || [];
        }

        const legacyMasterDataPermission = modulePermissions.find(mp => mp.moduleKey === 'masterdata' && mp.permissionLevel !== 'none');
        if (legacyMasterDataPermission) {
            modulePermissions = [
                ...modulePermissions,
                ...(modulePermissions.some(mp => mp.moduleKey === 'material') ? [] : [{
                    moduleKey: 'material' as ModuleKey,
                    permissionLevel: legacyMasterDataPermission.permissionLevel,
                }]),
                ...(modulePermissions.some(mp => mp.moduleKey === 'procurement') ? [] : [{
                    moduleKey: 'procurement' as ModuleKey,
                    permissionLevel: legacyMasterDataPermission.permissionLevel,
                }]),
            ];
        }

        // Build allowed modules and submodules based on permissions
        const permissionsMap = new Map<ModuleKey, PermissionLevel>();

        modulePermissions.forEach(mp => {
            if (mp.permissionLevel !== 'none') {
                allowedModules.push(mp.moduleKey);
                permissionsMap.set(mp.moduleKey, mp.permissionLevel);

                // Get submodules for this module
                const moduleDef = MODULE_DEFINITIONS.find(m => m.key === mp.moduleKey);
                if (moduleDef?.subModules) {
                    moduleDef.subModules.forEach(sub => {
                        if (hasPermissionLevel(mp.permissionLevel, sub.requiredPermission)) {
                            allowedSubModules.push(sub.key as SubModuleKey);
                        }
                    });
                }
            }
        });

        return {
            userId,
            userDisplayName,
            userEmail,
            role: roleKey,
            department,
            modules: allowedModules,
            subModules: allowedSubModules,
            permissions: permissionsMap,
            hasPermission: (moduleKey: ModuleKey, requiredLevel: PermissionLevel = 'view'): boolean => {
                const userLevel = permissionsMap.get(moduleKey);
                if (!userLevel) return false;
                return hasPermissionLevel(userLevel, requiredLevel);
            },
            hasSubModulePermission: (subModuleKey: SubModuleKey): boolean => {
                return allowedSubModules.includes(subModuleKey);
            },
        };
    }

    /**
     * Build default permissions for a given role
     */
    private buildDefaultPermissions(
        userId: number,
        userDisplayName: string,
        userEmail: string,
        roleKey: string
    ): IUserPermissions {
        return this.buildUserPermissions(
            userId,
            userDisplayName,
            userEmail,
            {
                Role: roleKey,
                Department: '',
                Modules: '',
            }
        );
    }

    /**
     * Clear cached permissions (useful when user role changes)
     */
    public clearCache(): void {
        this.cachedPermissions = null;
    }

    /**
     * Get all available roles
     */
    public getAvailableRoles(): IRolePermission[] {
        return DEFAULT_ROLE_PERMISSIONS;
    }

    /**
     * Get permissions for a specific role
     */
    public getRolePermissions(roleKey: string): IRolePermission | undefined {
        return DEFAULT_ROLE_PERMISSIONS.find(rp => rp.roleKey === roleKey);
    }

    /**
     * Check if current user has admin permissions
     */
    public async isAdmin(): Promise<boolean> {
        const permissions = await this.getUserPermissions();
        return permissions.role === 'Admin';
    }

    /**
     * Get all modules the current user can access
     */
    public async getAllowedModules(): Promise<ModuleKey[]> {
        const permissions = await this.getUserPermissions();
        return permissions.modules;
    }

    /**
     * Get all submodules the current user can access
     */
    public async getAllowedSubModules(): Promise<SubModuleKey[]> {
        const permissions = await this.getUserPermissions();
        return permissions.subModules;
    }
}

export default PermissionService;
