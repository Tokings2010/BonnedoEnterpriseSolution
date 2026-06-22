import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export const PROJECT_DOCUMENT_LIBRARY_TITLE = 'Documents';
export const PROJECT_DOCUMENT_LIBRARY_ROOT = 'Shared Documents';
export const PROJECT_BUDGET_FOLDER_NAME = 'Budgets';
export const PROJECT_WBS_FOLDER_NAME = 'Budgets';

export type ProjectDocumentKind = 'budget' | 'schedule';

export interface IProjectDocumentStorage {
  libraryTitle: string;
  libraryRootUrl: string;
  budgetFolderUrl: string;
  wbsFolderUrl: string;
  budgetFileName: string;
  wbsFileName: string;
}

export interface IStoredProjectDocument {
  name: string;
  serverRelativeUrl: string;
  url: string;
  length: number;
  modified: string;
}

interface IGetProjectDocumentLibraryRootUrlOptions {
  createIfMissing?: boolean;
}

const escapeODataString = (value: string): string => value.replace(/'/g, "''");

export const sanitizeProjectCodeForStorage = (projectCode: string): string =>
  projectCode
    .trim()
    .replace(/[^A-Za-z0-9._-]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '') || 'Project';

export const getProjectDocumentFileName = (projectCode: string, kind: ProjectDocumentKind): string =>
  `${projectCode}_${kind === 'budget' ? 'CBS' : 'WBS'}.xlsx`;

export const getProjectDocumentFolderServerRelativeUrl = (
  libraryRootUrl: string,
  folderName: string,
  projectCode: string
): string => `${libraryRootUrl.replace(/\/+$/, '')}/${folderName}/${projectCode}`;

export const getProjectDocumentServerRelativePath = (
  libraryRootUrl: string,
  folderName: string,
  projectCode: string,
  fileName: string
): string => `${getProjectDocumentFolderServerRelativeUrl(libraryRootUrl, folderName, projectCode)}/${fileName}`;

export const getProjectDocumentDisplayPath = (
  folderName: string,
  projectCode: string,
  fileName: string
): string => `${PROJECT_DOCUMENT_LIBRARY_TITLE}/${folderName}/${projectCode}/${fileName}`;

export const getProjectDocumentLibraryRootUrl = async (
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  options: IGetProjectDocumentLibraryRootUrlOptions = {}
): Promise<string> => {
  const webUrl = pageContext.web.absoluteUrl;
  const listUrl = `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(PROJECT_DOCUMENT_LIBRARY_TITLE)}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;

  const response: SPHttpClientResponse = await spHttpClient.get(listUrl, SPHttpClient.configurations.v1);

  if (response.ok) {
    const data = await response.json();
    return data.RootFolder?.ServerRelativeUrl || `${pageContext.web.serverRelativeUrl.replace(/\/+$/, '')}/${PROJECT_DOCUMENT_LIBRARY_ROOT}`;
  }

  if (!options.createIfMissing || response.status !== 404) {
    throw new Error(`Document library '${PROJECT_DOCUMENT_LIBRARY_TITLE}' was not found`);
  }

  const createListUrl = `${webUrl}/_api/web/lists`;
  const createListResponse: SPHttpClientResponse = await spHttpClient.post(
    createListUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
      },
      body: JSON.stringify({
        '__metadata': { type: 'SP.List' },
        BaseTemplate: 101,
        Title: PROJECT_DOCUMENT_LIBRARY_TITLE,
        Description: 'Project budget and schedule source files',
      }),
    }
  );

  if (!createListResponse.ok) {
    const errorText = await createListResponse.text().catch(() => '');
    throw new Error(`Failed to create document library '${PROJECT_DOCUMENT_LIBRARY_TITLE}': ${createListResponse.status} ${errorText}`);
  }

  return getProjectDocumentLibraryRootUrl(spHttpClient, pageContext, { createIfMissing: false });
};

const ensureFolderExists = async (
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  folderServerRelativeUrl: string
): Promise<void> => {
  const normalizedFolderUrl = folderServerRelativeUrl.startsWith('/')
    ? folderServerRelativeUrl
    : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
  const webUrl = pageContext.web.absoluteUrl;
  const folderUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(normalizedFolderUrl)}')`;

  const response: SPHttpClientResponse = await spHttpClient.get(folderUrl, SPHttpClient.configurations.v1);

  if (response.ok) {
    return;
  }

  const parentFolderUrl = normalizedFolderUrl.substring(0, normalizedFolderUrl.lastIndexOf('/'));
  const folderName = normalizedFolderUrl.substring(normalizedFolderUrl.lastIndexOf('/') + 1);

  if (!parentFolderUrl || !folderName) {
    throw new Error(`Invalid folder path: ${folderServerRelativeUrl}`);
  }

  await ensureFolderExists(spHttpClient, pageContext, parentFolderUrl);

  const createFolderUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(parentFolderUrl)}')/Folders/add('${encodeURIComponent(folderName)}')`;
  const createFolderResponse: SPHttpClientResponse = await spHttpClient.post(
    createFolderUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=verbose',
      },
      body: '',
    }
  );

  if (!createFolderResponse.ok) {
    const errorText = await createFolderResponse.text().catch(() => '');
    throw new Error(`Failed to create folder '${folderName}': ${createFolderResponse.status} ${errorText}`);
  }
};

export const getProjectDocumentStorage = async (
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  projectCode: string
): Promise<IProjectDocumentStorage> => {
  const storageProjectCode = sanitizeProjectCodeForStorage(projectCode);
  const libraryRootUrl = await getProjectDocumentLibraryRootUrl(spHttpClient, pageContext, { createIfMissing: true });

  const budgetFolderUrl = getProjectDocumentFolderServerRelativeUrl(libraryRootUrl, PROJECT_BUDGET_FOLDER_NAME, storageProjectCode);
  const wbsFolderUrl = getProjectDocumentFolderServerRelativeUrl(libraryRootUrl, PROJECT_WBS_FOLDER_NAME, storageProjectCode);

  await ensureFolderExists(spHttpClient, pageContext, getProjectDocumentFolderServerRelativeUrl(libraryRootUrl, PROJECT_BUDGET_FOLDER_NAME, storageProjectCode));
  await ensureFolderExists(spHttpClient, pageContext, budgetFolderUrl);
  await ensureFolderExists(spHttpClient, pageContext, getProjectDocumentFolderServerRelativeUrl(libraryRootUrl, PROJECT_WBS_FOLDER_NAME, storageProjectCode));
  await ensureFolderExists(spHttpClient, pageContext, wbsFolderUrl);

  return {
    libraryTitle: PROJECT_DOCUMENT_LIBRARY_TITLE,
    libraryRootUrl,
    budgetFolderUrl,
    wbsFolderUrl,
    budgetFileName: getProjectDocumentFileName(storageProjectCode, 'budget'),
    wbsFileName: getProjectDocumentFileName(storageProjectCode, 'schedule'),
  };
};

export const uploadProjectDocument = async (
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  file: File,
  fileName: string,
  folderServerRelativeUrl: string
): Promise<string> => {
  const normalizedFolderUrl = folderServerRelativeUrl.startsWith('/')
    ? folderServerRelativeUrl
    : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
  const uploadUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(normalizedFolderUrl)}')/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

  const uploadResponse: SPHttpClientResponse = await spHttpClient.post(
    uploadUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream',
      },
      body: await file.arrayBuffer(),
    }
  );

  if (!uploadResponse.ok) {
    const errorText = await uploadResponse.text().catch(() => '');
    throw new Error(`Upload failed for '${fileName}': ${uploadResponse.status} ${errorText}`);
  }

  const uploadData = await uploadResponse.json();
  return uploadData.d?.ServerRelativeUrl || uploadData.ServerRelativeUrl || `${folderServerRelativeUrl}/${fileName}`;
};

export const getStoredProjectDocument = async (
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  libraryRootUrl: string,
  folderName: string,
  projectCode: string,
  kind: ProjectDocumentKind
): Promise<IStoredProjectDocument | undefined> => {
  const storageProjectCode = sanitizeProjectCodeForStorage(projectCode);
  const fileName = getProjectDocumentFileName(storageProjectCode, kind);
  const fileServerRelativePath = getProjectDocumentServerRelativePath(libraryRootUrl, folderName, storageProjectCode, fileName);
  const fileUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativePath)}')?$select=Name,ServerRelativeUrl,Length,TimeLastModified`;

  try {
    const response: SPHttpClientResponse = await spHttpClient.get(fileUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      return undefined;
    }

    const data = await response.json();
    const serverRelativeUrl = data.ServerRelativeUrl || fileServerRelativePath;

    return {
      name: data.Name || fileName,
      serverRelativeUrl,
      url: `${window.location.origin}${serverRelativeUrl}`,
      length: data.Length || 0,
      modified: data.TimeLastModified || '',
    };
  } catch (error) {
    console.warn(`[ProjectDocumentStorage] Failed to read stored ${kind} document:`, error);
    return undefined;
  }
};
