import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

type MSGraphClient = any;

export interface ITaskRegisterItem {
  ID?: number;
  Title: string;
  Project_Code: string;
  Bucket?: string;
  AssignedToEmail?: string;
  StartDate?: string;
  DueDate?: string;
  Priority?: string;
  Status?: string;
  PlannerPlanId?: string;
  PlannerTaskId?: string;
  SyncLock?: boolean;
  LastSynced?: string;
}

export class TaskRegisterService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;
  private graphClient: MSGraphClient | null = null;
  private webPartContext: WebPartContext;

  private readonly LIST_NAME = 'Project_Task_Register';
  private readonly PLANNER_BUCKETS = ['Planning', 'Execution', 'Procurement', 'Finance', 'WCC and Invoicing', 'Closeout'];

  constructor(webPartContext: WebPartContext) {
    this.webPartContext = webPartContext;
    this.spHttpClient = webPartContext.spHttpClient;
    this.pageContext = webPartContext.pageContext;
  }

  private async initializeGraphClient(): Promise<MSGraphClient> {
    if (this.graphClient) return this.graphClient;
    this.graphClient = await this.webPartContext.msGraphClientFactory.getClient('3');
    return this.graphClient;
  }

  public async createTaskInSharePointAndPlanner(task: ITaskRegisterItem, projectPlanId?: string): Promise<ITaskRegisterItem> {
    // Create in SharePoint first
    const createdItem = await this.createSharePointTask(task);
    
    // Then sync to Planner
    if (projectPlanId || createdItem.PlannerPlanId) {
      const planId = projectPlanId || createdItem.PlannerPlanId!;
      const plannerTask = await this.createPlannerTask(planId, task);
      // Update SP with Planner IDs
      await this.updateSharePointWithPlannerIds(createdItem.ID!, plannerTask.id, planId);
      createdItem.PlannerTaskId = plannerTask.id;
      createdItem.PlannerPlanId = planId;
    }
    
    return createdItem;
  }

  private async createSharePointTask(task: ITaskRegisterItem): Promise<ITaskRegisterItem> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items`;
    
    const body: any = {
      Title: task.Title,
      Project_Code: task.Project_Code,
      Bucket: task.Bucket || 'Planning',
      AssignedToEmail: task.AssignedToEmail,
      StartDate: task.StartDate,
      DueDate: task.DueDate,
      Priority: task.Priority || 'High',
      Status: task.Status || 'Not Started',
      SyncLock: false
    };

    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    if (!response.ok) throw new Error('Failed to create task in SharePoint');
    // SharePoint POST requests return the created item in the response body
    return await response.json();
  }

  private async createPlannerTask(planId: string, task: ITaskRegisterItem): Promise<{ id: string }> {
    const graphClient = await this.initializeGraphClient();
    const dueDate = task.DueDate ? new Date(task.DueDate).toISOString() : undefined;
    
    const payload: any = {
      planId: planId,
      title: task.Title,
      bucketId: await this.getBucketId(planId, task.Bucket || 'Planning'),
      assignments: (task as any).AssignedToId ? await this.buildAssignments((task as any).AssignedToId) : undefined,
      dueDateTime: dueDate
    };

    const response = await graphClient.api('planner/tasks').post(payload);
    return { id: response.id };
  }

  private async getBucketId(planId: string, bucketName: string): Promise<string | undefined> {
    const graphClient = await this.initializeGraphClient();
    const bucketsResponse = await graphClient.api(`planner/plans/${planId}/buckets`).get();
    const bucket = bucketsResponse.value.find((b: any) => b.name === bucketName);
    return bucket?.id;
  }

  private async buildAssignments(userId: string): Promise<any> {
    if (!userId) return undefined;
    return {
      [userId]: {
        '@odata.type': 'microsoft.graph.plannerAssignment',
        orderHint: ' !'
      }
    };
  }

  private async updateSharePointWithPlannerIds(itemId: number, plannerTaskId: string, planId: string): Promise<void> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const getItemUrl = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${itemId})`;

    // First, get the current item to get the etag from response headers
    const getItemResponse = await this.spHttpClient.get(getItemUrl, SPHttpClient.configurations.v1);
    if (!getItemResponse.ok) {
      throw new Error('Failed to retrieve task for etag');
    }
    // Get etag from response headers (case-insensitive), fallback to '*' if not available
    const etag = getItemResponse.headers.get('etag') || getItemResponse.headers.get('ETag') || '*';

    const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${itemId})`;
    
    await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE',
        'If-Match': etag
      },
      body: JSON.stringify({
        PlannerTaskId: plannerTaskId,
        PlannerPlanId: planId,
        LastSynced: new Date().toISOString()
      })
    });
  }

  // Reverse sync - poll or webhook based update from Planner to SP
  public async syncPlannerTaskToSharePoint(plannerTaskId: string, updates: Partial<ITaskRegisterItem>): Promise<void> {
    // Find SP item by PlannerTaskId
    const item = await this.getTaskByPlannerId(plannerTaskId);
    if (!item) return;

    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${item.ID})`;

    const body: any = {};
    if (updates.Status) body.Status = updates.Status;
    if (updates.Priority) body.Priority = updates.Priority;
    body.LastSynced = new Date().toISOString();

    await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json', 'Content-Type': 'application/json', 'X-HTTP-Method': 'MERGE', 'If-Match': '*' },
      body: JSON.stringify(body)
    });
  }

  private async getTaskByPlannerId(plannerTaskId: string): Promise<ITaskRegisterItem | null> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items?$filter=PlannerTaskId eq '${plannerTaskId}'&$top=1`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return null;
    const data = await response.json();
    return data.value?.[0] || null;
  }

  public async getProjectPlanId(projectCode: string): Promise<string | null> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$filter=Project_Code eq '${projectCode}'&$select=Planner_Plan_ID,Teams_Group_ID&$top=1`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return null;
    const data = await response.json();
    const item = data.value?.[0];
    if (item) {
      return item.Planner_Plan_ID || null;
    }
    return null;
  }

  public async getProjectDetails(projectCode: string): Promise<{ plannerPlanId: string | null; teamsGroupId: string | null } | null> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const url = `${webUrl}/_api/web/lists/getByTitle('ENT_Project_Master')/items?$filter=Project_Code eq '${projectCode}'&$select=Planner_Plan_ID,Teams_Group_ID&$top=1`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return null;
    const data = await response.json();
    const item = data.value?.[0];
    if (item) {
      return {
        plannerPlanId: item.Planner_Plan_ID || null,
        teamsGroupId: item.Teams_Group_ID || null
      };
    }
    return null;
  }

   public async getTeamMembers(groupId: string): Promise<Array<{ id: string; displayName: string; email: string }>> {
      if (!groupId) return [];
      const graphClient = await this.initializeGraphClient();
      try {
        const response = await graphClient.api(`groups/${groupId}/members`).select('id,displayName,mail').get();
        return response.value?.map((m: any) => ({
          id: m.id,
          displayName: m.displayName,
          email: m.mail || ''
        })) || [];
      } catch (e) {
        console.warn('Failed to fetch team members', e);
        return [];
      }
   }

    /**
    * Update a task in SharePoint
    * @param task The task item to update (must include ID)
    */
    private async updateSharePointTask(task: ITaskRegisterItem): Promise<ITaskRegisterItem> {
      if (!task.ID) {
        throw new Error('Task ID is required to update a SharePoint task');
      }

      const webUrl = this.pageContext.web.absoluteUrl;
      const getItemUrl = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${task.ID})`;

      // First, get the current item to get the etag from headers
      const getItemResponse = await this.spHttpClient.get(getItemUrl, SPHttpClient.configurations.v1);
      if (!getItemResponse.ok) {
        throw new Error('Failed to retrieve task for etag');
      }
      // Get etag from response headers (case-insensitive)
      const etag = getItemResponse.headers.get('etag') || getItemResponse.headers.get('ETag') || '*';

      const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${task.ID})`;

      const body: any = {
        Title: task.Title,
        Project_Code: task.Project_Code,
        Bucket: task.Bucket || 'Planning',
        AssignedToEmail: task.AssignedToEmail,
        StartDate: task.StartDate,
        DueDate: task.DueDate,
        Priority: task.Priority || 'High',
        Status: task.Status || 'Not Started'
      };

      const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-HTTP-Method': 'MERGE',
          'If-Match': etag
        },
        body: JSON.stringify(body)
      });

      if (!response.ok) throw new Error('Failed to update task in SharePoint');
      // For MERGE requests, SharePoint returns 204 No Content with no body
      if (response.status === 204) {
        return task;
      }
      const text = await response.text();
      return text ? JSON.parse(text) : task;
    }

   /**
    * Update a task in Planner
    * @param plannerTaskId The Planner task ID
    * @param task The task data to update
    */
    private async updatePlannerTask(plannerTaskId: string, task: ITaskRegisterItem): Promise<void> {
      if (!task.PlannerPlanId) {
        throw new Error('PlannerPlanId is required to update a Planner task');
      }

      const graphClient = await this.initializeGraphClient();

      // First, GET the current Planner task to retrieve its @odata.etag for optimistic concurrency
      const currentTask = await graphClient.api(`planner/tasks/${plannerTaskId}`).get();
      const plannerEtag = currentTask['@odata.etag'];

      const dueDate = task.DueDate ? new Date(task.DueDate).toISOString() : undefined;

      // Get the bucket ID for the specified bucket name in the plan
      const bucketId = await this.getBucketId(task.PlannerPlanId, task.Bucket || 'Planning');

      // Map SharePoint Status to Planner percentComplete
      let percentComplete = 0;
      if (task.Status === 'Completed') {
        percentComplete = 100;
      } else if (task.Status === 'In Progress') {
        percentComplete = 50;
      } else {
        percentComplete = 0; // Not Started or default
      }

      const payload: any = {
        title: task.Title,
        bucketId: bucketId,
        dueDateTime: dueDate,
        percentComplete: percentComplete
      };

      // Handle assignments if AssignedToId is present (we assume it's stored in task during edit)
      if ((task as any).AssignedToId) {
        payload.assignments = await this.buildAssignments((task as any).AssignedToId);
      }

      // Update with If-Match header containing the current etag to satisfy Planner's concurrency requirements
      await graphClient
        .api(`planner/tasks/${plannerTaskId}`)
        .header('If-Match', plannerEtag)
        .update(payload);
    }

   /**
    * Update a task in both SharePoint and Planner
    * @param task The task item to update (must include ID)
    */
   public async updateTaskInSharePointAndPlanner(task: ITaskRegisterItem): Promise<ITaskRegisterItem> {
     // Update SharePoint first
     const updatedItem = await this.updateSharePointTask(task);

     // If we have a PlannerTaskId, update the Planner task
     if (updatedItem.PlannerTaskId) {
       await this.updatePlannerTask(updatedItem.PlannerTaskId, task);
     }

     return updatedItem;
   }

   /**
    * Sync the status of a Planner task back to SharePoint
    * @param plannerTaskId The Planner task ID
    */
   public async syncPlannerTaskStatusToSharePoint(plannerTaskId: string): Promise<void> {
     try {
       // 1. Get the task from Planner
       const graphClient = await this.initializeGraphClient();
       const plannerTask = await graphClient.api(`planner/tasks/${plannerTaskId}`).get();

       // 2. Extract the status (percentComplete) and map to our status
       // Planner uses percentComplete: 0 (Not Started), 1-99 (In Progress), 100 (Completed)
       let status: string = 'Not Started';
       const percentComplete = plannerTask.percentComplete || 0;
       if (percentComplete === 100) {
         status = 'Completed';
       } else if (percentComplete > 0) {
         status = 'In Progress';
       }

       // 3. Find the SharePoint item by PlannerTaskId
       const item = await this.getTaskByPlannerId(plannerTaskId);
       if (!item) {
         console.warn(`No SharePoint item found for Planner task ID: ${plannerTaskId}`);
         return;
       }

       // 4. Update the status in SharePoint
       const webUrl = this.pageContext.web.absoluteUrl;
       const url = `${webUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/items(${item.ID})`;

       await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
         headers: {
           'Accept': 'application/json',
           'Content-Type': 'application/json',
           'X-HTTP-Method': 'MERGE',
           'If-Match': '*'
         },
         body: JSON.stringify({
           Status: status,
           LastSynced: new Date().toISOString()
         })
       });
     } catch (error) {
       console.error('Error syncing Planner task status to SharePoint:', error);
       throw error;
     }
   }
}

export function createTaskRegisterService(webPartContext: WebPartContext): TaskRegisterService {
  return new TaskRegisterService(webPartContext);
}