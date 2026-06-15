import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// MSGraphClientV3 type definition - it's obtained from the WebPart context
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClient = any;

// ============================================
// Project Provisioning Interfaces
// ============================================

export interface IProjectProvisioningRequest {
    projectId: number;
    projectName: string;
    projectCode: string;
    projectManagerEmail?: string;
}

export interface IProvisioningResult {
    success: boolean;
    teamId?: string;
    plannerId?: string;
    siteUrl?: string;
    error?: string;
}

export interface IProvisioningProgress {
    step: string;
    message: string;
    progress: number; // 0-100
}

// ============================================
// Project Provisioning Service
// ============================================

export class ProjectProvisioningService {
    private spHttpClient: SPHttpClient;
    private pageContext: PageContext;
    private graphClient: MSGraphClient | null = null;
    private webPartContext: WebPartContext;

    // Constants
    private readonly PROJECT_LIST_NAME = 'ENT_Project_Master';
    private readonly DEFAULT_CHANNELS = [
        'Planning',
        'Execution',
        'Procurement',
        'Finance',
        'WCC and Invoicing',
        'Closeout'
    ];
    private readonly PLANNER_BUCKETS = [
        'Contract & Mobilization',
        'Planning',
        'Execution',
        'WCC Certification',
        'Invoicing',
        'Closeout'
    ];
    private readonly DOCUMENT_TABS = [
        'Project Documents',
        'Procurement Docs',
        'Finance Docs'
    ];

    constructor(webPartContext: WebPartContext) {
        this.webPartContext = webPartContext;
        this.spHttpClient = webPartContext.spHttpClient;
        this.pageContext = webPartContext.pageContext;
    }

    // ============================================
    // Initialize Graph Client
    // ============================================

    private async initializeGraphClient(): Promise<MSGraphClient> {
        if (this.graphClient) {
            return this.graphClient;
        }

        // Get MSGraphClientV3 from the WebPart context
        this.graphClient = await this.webPartContext.msGraphClientFactory.getClient('3');

        return this.graphClient;
    }

    // ============================================
    // STEP 13: Expose provisionProjectWorkspace function
    // ============================================

    /**
     * Main function to provision project workspace
     * This function will be called after project creation
     * @param request - Project provisioning request parameters
     * @param onProgress - Optional callback for progress updates
     * @returns Promise with provisioning result
     */
    public async provisionProjectWorkspace(
        request: IProjectProvisioningRequest,
        onProgress?: (progress: IProvisioningProgress) => void
    ): Promise<IProvisioningResult> {
        console.log('[ProjectProvisioning] Starting project workspace provisioning:', request);

        let graphClient: MSGraphClient;

        try {
            // STEP 14: Ensure idempotency - check if Teams_Group_ID already exists
            onProgress?.({ step: 'check', message: 'Checking existing workspace...', progress: 5 });
            const existingProject = await this.getProjectById(request.projectId);
            if (existingProject && existingProject.Teams_Group_ID) {
                console.log('[ProjectProvisioning] Project already has Teams_Group_ID, skipping provisioning');
                onProgress?.({ step: 'complete', message: 'Workspace already exists', progress: 100 });
                return {
                    success: true,
                    teamId: existingProject.Teams_Group_ID,
                    plannerId: existingProject.Planner_Plan_ID,
                    siteUrl: existingProject.Teams_Site_URL
                };
            }

            // Initialize Graph client
            onProgress?.({ step: 'init', message: 'Initializing...', progress: 10 });
            graphClient = await this.initializeGraphClient();

            // STEP 2: Update SharePoint project record - Set Provisioning_Status = "Provisioning"
            onProgress?.({ step: 'status', message: 'Updating provisioning status...', progress: 15 });
            await this.updateProvisioningStatus(request.projectId, 'Provisioning');

            // STEP 3: Create Microsoft Team
            onProgress?.({ step: 'team', message: 'Creating Microsoft Team...', progress: 20 });
            console.log('[ProjectProvisioning] Creating team with name:', `Project - ${request.projectName}`);
            const teamResponse = await this.createTeam(graphClient, request.projectName);
            const groupId = teamResponse.id;
            console.log('[ProjectProvisioning] Team created with Group ID:', groupId);

            if (!groupId) {
                throw new Error('Failed to get Group ID from team creation response');
            }

            // STEP 4: Update SharePoint project record - Save Teams_Group_ID
            onProgress?.({ step: 'teamId', message: 'Saving team information...', progress: 35 });
            await this.updateTeamGroupId(request.projectId, groupId);

            // Wait for team to be fully provisioned
            onProgress?.({ step: 'provisioning', message: 'Waiting for team provisioning...', progress: 40 });
            await this.waitForTeamProvisioning(graphClient, groupId);

            // STEP 5: Create default channels
            onProgress?.({ step: 'channels', message: 'Creating channels...', progress: 50 });
            const channels = await this.createDefaultChannels(graphClient, groupId);
            console.log('[ProjectProvisioning] Default channels created');

            // STEP 6: Create Project Folder inside General channel document library
            onProgress?.({ step: 'folder', message: 'Creating project folder...', progress: 60 });
            const projectFolderUrl = await this.createProjectFolder(graphClient, groupId, request.projectName);
            console.log('[ProjectProvisioning] Project folder created:', projectFolderUrl);

            // STEP 7: Create Planner Plan
            onProgress?.({ step: 'planner', message: 'Creating planner plan...', progress: 70 });
            const plannerPlan = await this.createPlannerPlan(graphClient, groupId, request.projectName);
            const planId = plannerPlan.id;
            console.log('[ProjectProvisioning] Planner plan created with ID:', planId);

            // STEP 8: Create default Planner buckets
            onProgress?.({ step: 'buckets', message: 'Creating planner buckets...', progress: 80 });
            await this.createPlannerBuckets(graphClient, planId);
            console.log('[ProjectProvisioning] Planner buckets created');

            // STEP 9: Pin Planner tab in General channel
            onProgress?.({ step: 'tab', message: 'Adding planner tab...', progress: 85 });
            const generalChannelId = channels.find(c => c.displayName === 'General')?.id;
            if (generalChannelId) {
                await this.addPlannerTab(graphClient, groupId, generalChannelId);
                console.log('[ProjectProvisioning] Planner tab added to General channel');
            }

            // STEP 10: Add Document Library tab to channels
            // Note: Document library tabs are currently not supported via Graph API
            // They need to be added manually in Teams
            console.log('[ProjectProvisioning] Skipping document library tabs (not supported via Graph API)');

            // STEP 11: Update SharePoint project record
            // Save: Planner_Plan_ID, Project_Folder_URL, Teams_Site_URL
            // Set: Provisioning_Status = "Completed"
            onProgress?.({ step: 'sharepoint', message: 'Updating SharePoint...', progress: 90 });
            const siteUrl = `https://teams.microsoft.com/l/team/${groupId}/conversations`;
            await this.updateProjectWithCompletion(
                request.projectId,
                planId,
                projectFolderUrl,
                siteUrl
            );

            console.log('[ProjectProvisioning] Project provisioning completed successfully');
            onProgress?.({ step: 'complete', message: 'Provisioning completed!', progress: 100 });

            return {
                success: true,
                teamId: groupId,
                plannerId: planId,
                siteUrl: siteUrl
            };

        } catch (error) {
            // STEP 12: Implement error handling
            console.error('[ProjectProvisioning] Error during provisioning:', error);
            onProgress?.({ step: 'error', message: 'Provisioning failed', progress: 0 });

            try {
                await this.updateProvisioningStatus(request.projectId, 'Failed');
                console.log('[ProjectProvisioning] Updated Provisioning_Status to Failed');
            } catch (updateError) {
                console.error('[ProjectProvisioning] Failed to update status to Failed:', updateError);
            }

            return {
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error occurred'
            };
        }
    }

    // ============================================
    // Helper Methods
    // ============================================

    /**
     * Get project by ID
     */
    private async getProjectById(projectId: number): Promise<{ Teams_Group_ID?: string; Planner_Plan_ID?: string; Teams_Site_URL?: string } | null> {
        const webUrl = this.pageContext.web.absoluteUrl;
        const url = `${webUrl}/_api/web/lists/getByTitle('${this.PROJECT_LIST_NAME}')/items(${projectId})?$select=Teams_Group_ID,Planner_Plan_ID,Teams_Site_URL`;

        try {
            const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!response.ok) return null;
            return await response.json();
        } catch {
            return null;
        }
    }

    /**
     * Update Provisioning Status
     * STEP 2 & STEP 12: Set Provisioning_Status to "Provisioning" or "Failed"
     */
    private async updateProvisioningStatus(projectId: number, status: string): Promise<void> {
        const webUrl = this.pageContext.web.absoluteUrl;
        const url = `${webUrl}/_api/web/lists/getByTitle('${this.PROJECT_LIST_NAME}')/items(${projectId})`;

        const response = await this.spHttpClient.post(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*'
                },
                body: JSON.stringify({
                    Provisioning_Status: status
                })
            }
        );

        if (!response.ok) {
            const errorText = await response.text();
            console.error('[ProjectProvisioning] Failed to update provisioning status:', response.status, errorText);
            throw new Error(`Failed to update provisioning status: ${response.status}`);
        }
    }

    /**
     * STEP 4: Update SharePoint project record - Save Teams_Group_ID
     */
    private async updateTeamGroupId(projectId: number, groupId: string): Promise<void> {
        const webUrl = this.pageContext.web.absoluteUrl;
        const url = `${webUrl}/_api/web/lists/getByTitle('${this.PROJECT_LIST_NAME}')/items(${projectId})`;

        const response = await this.spHttpClient.post(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*'
                },
                body: JSON.stringify({
                    Teams_Group_ID: groupId
                })
            }
        );

        if (!response.ok) {
            const errorText = await response.text();
            console.error('[ProjectProvisioning] Failed to update team group ID:', response.status, errorText);
            throw new Error(`Failed to update team group ID: ${response.status}`);
        }
    }

    /**
     * STEP 3: Create Microsoft Team using Graph API
     * POST /teams
     * 
     * IMPORTANT: The Graph API returns a 202 Accepted status with a Location header
     * when creating teams. The response body is a ReadableStream, not JSON.
     * Since MSGraphClientV3 doesn't expose response headers directly,
     * we query the Graph API to find the newly created team by name.
     */
    private async createTeam(graphClient: MSGraphClient, projectName: string): Promise<{ id: string }> {
        const teamDisplayName = `Project - ${projectName}`;
        const teamPayload = {
            'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'standard\')',
            displayName: teamDisplayName,
            description: `Project workspace for ${projectName}`
        };

        try {
            console.log('[ProjectProvisioning] Sending team creation request with payload:', JSON.stringify(teamPayload));

            // Call the Graph API - this returns 202 Accepted with empty body
            const response = await graphClient.api('teams').post(teamPayload);

            console.log('[ProjectProvisioning] Raw response type:', typeof response);
            console.log('[ProjectProvisioning] Raw response value:', response);

            // The Graph API returns 202 Accepted with empty body
            // We need to query for the team by name to get its ID
            console.log('[ProjectProvisioning] Team creation request accepted (202). Querying for team by name...');

            // Wait a moment for the team to be created
            await new Promise(resolve => setTimeout(resolve, 3000));

            // Query for the team by display name
            let groupId: string | undefined;
            let attempts = 0;
            const maxAttempts = 10;

            while (!groupId && attempts < maxAttempts) {
                attempts++;
                console.log(`[ProjectProvisioning] Attempt ${attempts}/${maxAttempts} to find team by name`);

                try {
                    // Search for groups with matching display name
                    // Use $filter to find groups created recently with matching name
                    const filterQuery = `displayName eq '${teamDisplayName.replace(/'/g, '\'\'')}'`;
                    const groupsResponse = await graphClient.api('groups')
                        .filter(filterQuery)
                        .select('id,displayName,resourceProvisioningOptions')
                        .get();

                    console.log('[ProjectProvisioning] Groups query response:', groupsResponse);

                    if (groupsResponse && groupsResponse.value && groupsResponse.value.length > 0) {
                        // Find the team (group with resourceProvisioningOptions containing 'Team')
                        const teamGroup = groupsResponse.value.find((g: any) =>
                            g.resourceProvisioningOptions &&
                            g.resourceProvisioningOptions.includes('Team')
                        );

                        if (teamGroup) {
                            groupId = teamGroup.id;
                            console.log('[ProjectProvisioning] Found team with ID:', groupId);
                            break;
                        } else {
                            console.log('[ProjectProvisioning] Found groups but none are teams yet, waiting...');
                        }
                    } else {
                        console.log('[ProjectProvisioning] No groups found yet, waiting...');
                    }
                } catch (queryError) {
                    console.warn('[ProjectProvisioning] Error querying for team:', queryError);
                }

                // Wait 2 seconds before next attempt
                if (!groupId && attempts < maxAttempts) {
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }

            if (!groupId) {
                console.error('[ProjectProvisioning] Could not find team after multiple attempts');
                throw new Error('Team creation failed - could not find team by name. The team may still be provisioning. Please check Microsoft Teams admin center.');
            }

            console.log('[ProjectProvisioning] Successfully created team with ID:', groupId);
            return { id: groupId };
        } catch (error) {
            console.error('[ProjectProvisioning] Error creating team:', error);

            let errorMessage = 'Unknown error';
            if (error instanceof Error) {
                errorMessage = error.message;
            }

            throw new Error(`Failed to create team: ${errorMessage}`);
        }
    }

    /**
     * Wait for team to be fully provisioned
     */
    private async waitForTeamProvisioning(graphClient: MSGraphClient, groupId: string, maxAttempts: number = 30): Promise<void> {
        console.log(`[ProjectProvisioning] Waiting for team provisioning, Group ID: ${groupId}`);

        for (let i = 0; i < maxAttempts; i++) {
            try {
                console.log(`[ProjectProvisioning] Attempt ${i + 1}/${maxAttempts} to check team status`);
                const response = await graphClient.api(`teams/${groupId}`).get();
                console.log('[ProjectProvisioning] Team status response:', response);

                if (response && response.id) {
                    console.log('[ProjectProvisioning] Team provisioned successfully');
                    return;
                }
            } catch (error) {
                // Team might not be ready yet, continue waiting
                console.warn(`[ProjectProvisioning] Team not ready yet (attempt ${i + 1}):`, error instanceof Error ? error.message : 'Unknown error');
            }

            // Wait 2 seconds between attempts
            await new Promise(resolve => setTimeout(resolve, 2000));
        }

        console.error('[ProjectProvisioning] Team provisioning timeout after max attempts');
        throw new Error('Team provisioning timeout - Team creation may have failed. Please check Microsoft 365 admin center for team status.');
    }

    /**
     * STEP 5: Create default channels
     * POST /teams/{team-id}/channels
     */
    private async createDefaultChannels(graphClient: MSGraphClient, groupId: string): Promise<Array<{ id: string; displayName: string }>> {
        const channels: Array<{ id: string; displayName: string }> = [];

        for (const channelName of this.DEFAULT_CHANNELS) {
            try {
                const channelPayload = {
                    displayName: channelName,
                    description: `${channelName} channel for project`
                };

                const response = await graphClient.api(`teams/${groupId}/channels`).post(channelPayload);
                channels.push({ id: response.id, displayName: response.displayName });
            } catch (error) {
                console.warn(`Failed to create channel ${channelName}, continuing...`, error);
            }
        }

        return channels;
    }

    /**
     * STEP 6: Create Project Folder inside General channel document library
     * POST /sites/{site-id}/drive/root/children
     */
    private async createProjectFolder(graphClient: MSGraphClient, groupId: string, projectName: string): Promise<string> {
        // First, get the team site ID
        const siteResponse = await graphClient.api(`groups/${groupId}/sites/root`).get();
        const siteId = siteResponse.id;

        // Create folder in the root of the document library
        const folderPayload = {
            name: projectName,
            folder: {}
        };

        const response = await graphClient.api(`sites/${siteId}/drive/root/children`).post(folderPayload);

        return response.webUrl;
    }

    /**
     * STEP 7: Create Planner Plan
     * POST /planner/plans
     */
    private async createPlannerPlan(graphClient: MSGraphClient, groupId: string, projectName: string): Promise<{ id: string }> {
        const planPayload = {
            owner: groupId,
            title: `Project Plan - ${projectName}`
        };

        const response = await graphClient.api('planner/plans').post(planPayload);

        return { id: response.id };
    }

    /**
     * STEP 8: Create default Planner buckets
     * POST /planner/buckets
     */
    private async createPlannerBuckets(graphClient: MSGraphClient, planId: string): Promise<void> {
        for (const bucketName of this.PLANNER_BUCKETS) {
            try {
                const bucketPayload = {
                    name: bucketName,
                    planId: planId
                };

                await graphClient.api('planner/buckets').post(bucketPayload);
            } catch (error) {
                console.warn(`Failed to create bucket ${bucketName}, continuing...`, error);
            }
        }
    }

    /**
     * STEP 9: Pin Planner tab in General channel
     * POST /teams/{team-id}/channels/{channel-id}/tabs
     */
    private async addPlannerTab(graphClient: MSGraphClient, teamId: string, channelId: string): Promise<void> {
        // The Planner app ID in Teams - using the enterprise ID
        const plannerAppId = 'com.microsoft.teamspace.tab.planner';

        const tabPayload = {
            displayName: 'Project Planner',
            'teamsApp@odata.bind': `https://graph.microsoft.com/v1.0/teamsApps/${plannerAppId}`,
            configuration: {
                entityId: '',
                contentUrl: '',
                websiteUrl: ''
            }
        };

        try {
            await graphClient.api(`teams/${teamId}/channels/${channelId}/tabs`).post(tabPayload);
        } catch (error) {
            console.warn('Failed to add Planner tab, continuing...', error);
        }
    }

    /**
     * STEP 10: Add Document Library tab to channels
     * POST /teams/{team-id}/channels/{channel-id}/tabs
     * 
     * Note: Document library tabs use @microsoft.graph.templateId
     */
    private async addDocumentLibraryTabs(
        graphClient: MSGraphClient,
        teamId: string,
        channels: Array<{ id: string; displayName: string }>
    ): Promise<void> {
        // Map of channel names to document tabs
        const channelTabMapping: { [key: string]: string[] } = {
            'General': ['Project Documents'],
            'Procurement': ['Procurement Docs'],
            'Finance': ['Finance Docs']
        };

        for (const channel of channels) {
            const tabsToAdd = channelTabMapping[channel.displayName] || [];

            for (const tabName of tabsToAdd) {
                try {
                    const tabPayload = {
                        displayName: tabName,
                        '@microsoft.graph.templateId': 'com.microsoft.teamspace.tab.files',
                        configuration: {
                            entityId: '',
                            contentUrl: '',
                            websiteUrl: '',
                            removeUrl: ''
                        }
                    };

                    await graphClient.api(`teams/${teamId}/channels/${channel.id}/tabs`).post(tabPayload);
                    console.log(`[ProjectProvisioning] Added ${tabName} tab to ${channel.displayName}`);
                } catch (error) {
                    console.warn(`Failed to add ${tabName} tab to ${channel.displayName}, continuing...`, error);
                }
            }
        }
    }

    /**
     * STEP 11: Update SharePoint project record with completion data
     * Save: Planner_Plan_ID, Project_Folder_URL, Teams_Site_URL
     * Set: Provisioning_Status = "Completed"
     * 
     * IMPORTANT: Project_Folder_URL and Teams_Site_URL are Hyperlink/URL fields in SharePoint
     * which require a special JSON format with both Url and Description properties.
     */
    private async updateProjectWithCompletion(
        projectId: number,
        planId: string,
        projectFolderUrl: string,
        teamsSiteUrl: string
    ): Promise<void> {
        const webUrl = this.pageContext.web.absoluteUrl;
        const url = `${webUrl}/_api/web/lists/getByTitle('${this.PROJECT_LIST_NAME}')/items(${projectId})`;

        console.log('[ProjectProvisioning] Updating SharePoint with completion data:', {
            projectId,
            planId,
            projectFolderUrl,
            teamsSiteUrl
        });

        // Format URL fields for SharePoint Hyperlink/URL column type
        // SharePoint requires { Url: string, Description: string } format
        const formattedProjectFolderUrl = {
            Url: projectFolderUrl,
            Description: 'Project Folder'
        };

        const formattedTeamsSiteUrl = {
            Url: teamsSiteUrl,
            Description: 'Teams Site'
        };

        try {
            const response = await this.spHttpClient.post(
                url,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-HTTP-Method': 'MERGE',
                        'If-Match': '*'
                    },
                    body: JSON.stringify({
                        Planner_Plan_ID: planId,
                        Project_Folder_URL: formattedProjectFolderUrl,
                        Teams_Site_URL: formattedTeamsSiteUrl,
                        Provisioning_Status: 'Completed'
                    })
                }
            );

            // Check if the response was successful
            if (!response.ok) {
                const errorText = await response.text();
                console.error('[ProjectProvisioning] SharePoint update failed:', response.status, errorText);
                throw new Error(`SharePoint update failed with status ${response.status}: ${errorText}`);
            }

            console.log('[ProjectProvisioning] SharePoint update completed successfully');
        } catch (error) {
            console.error('[ProjectProvisioning] Error updating SharePoint:', error);
            // Try to update fields individually if the full update fails
            try {
                console.log('[ProjectProvisioning] Attempting individual field updates...');

                // Update Provisioning_Status first
                const statusResponse = await this.spHttpClient.post(
                    url,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-HTTP-Method': 'MERGE',
                            'If-Match': '*'
                        },
                        body: JSON.stringify({
                            Provisioning_Status: 'Completed'
                        })
                    }
                );

                if (!statusResponse.ok) {
                    const errorText = await statusResponse.text();
                    console.warn('[ProjectProvisioning] Failed to update Provisioning_Status:', statusResponse.status, errorText);
                } else {
                    console.log('[ProjectProvisioning] Provisioning_Status updated');
                }

                // Update Planner_Plan_ID
                try {
                    const plannerResponse = await this.spHttpClient.post(
                        url,
                        SPHttpClient.configurations.v1,
                        {
                            headers: {
                                'Accept': 'application/json',
                                'Content-Type': 'application/json',
                                'X-HTTP-Method': 'MERGE',
                                'If-Match': '*'
                            },
                            body: JSON.stringify({
                                Planner_Plan_ID: planId
                            })
                        }
                    );

                    if (!plannerResponse.ok) {
                        const errorText = await plannerResponse.text();
                        console.warn('[ProjectProvisioning] Failed to update Planner_Plan_ID:', plannerResponse.status, errorText);
                    } else {
                        console.log('[ProjectProvisioning] Planner_Plan_ID updated');
                    }
                } catch (plannerError) {
                    console.warn('[ProjectProvisioning] Failed to update Planner_Plan_ID:', plannerError);
                }

                // Update Project_Folder_URL (with proper Hyperlink field format)
                try {
                    const folderResponse = await this.spHttpClient.post(
                        url,
                        SPHttpClient.configurations.v1,
                        {
                            headers: {
                                'Accept': 'application/json',
                                'Content-Type': 'application/json',
                                'X-HTTP-Method': 'MERGE',
                                'If-Match': '*'
                            },
                            body: JSON.stringify({
                                Project_Folder_URL: formattedProjectFolderUrl
                            })
                        }
                    );

                    if (!folderResponse.ok) {
                        const errorText = await folderResponse.text();
                        console.warn('[ProjectProvisioning] Failed to update Project_Folder_URL:', folderResponse.status, errorText);
                    } else {
                        console.log('[ProjectProvisioning] Project_Folder_URL updated');
                    }
                } catch (folderError) {
                    console.warn('[ProjectProvisioning] Failed to update Project_Folder_URL:', folderError);
                }

                // Update Teams_Site_URL (with proper Hyperlink field format)
                try {
                    const siteResponse = await this.spHttpClient.post(
                        url,
                        SPHttpClient.configurations.v1,
                        {
                            headers: {
                                'Accept': 'application/json',
                                'Content-Type': 'application/json',
                                'X-HTTP-Method': 'MERGE',
                                'If-Match': '*'
                            },
                            body: JSON.stringify({
                                Teams_Site_URL: formattedTeamsSiteUrl
                            })
                        }
                    );

                    if (!siteResponse.ok) {
                        const errorText = await siteResponse.text();
                        console.warn('[ProjectProvisioning] Failed to update Teams_Site_URL:', siteResponse.status, errorText);
                    } else {
                        console.log('[ProjectProvisioning] Teams_Site_URL updated');
                    }
                } catch (siteError) {
                    console.warn('[ProjectProvisioning] Failed to update Teams_Site_URL:', siteError);
                }

                console.log('[ProjectProvisioning] Individual field updates completed');
            } catch (individualError) {
                console.error('[ProjectProvisioning] Individual field updates also failed:', individualError);
                throw error; // Throw the original error
            }
        }
    }
}

// ============================================
// Factory function to create service instance
// ============================================

export function createProjectProvisioningService(webPartContext: WebPartContext): ProjectProvisioningService {
    return new ProjectProvisioningService(webPartContext);
}