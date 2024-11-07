import fs from 'fs';

interface accessTokenResponseType {  
  /**
   * The requested access token. Your app can use this token in calls to Microsoft Graph.
   */
  access_token: string;
  /**
   * How long the access token is valid (in seconds).
   */
  expires_in: number;
  /**
   * Used to indicate an extended lifetime for the access token and to support resiliency when the token issuance service isn't responding.
   */

  ext_expires_in: number;
  /**
   * Indicates the token type value. The only type that Microsoft Entra ID supports is Bearer.
   */
  token_type: string;
}

export interface FolderDetails extends FileOrFolderDetails {
  folder: object;
}

export interface FileOrFolderDetails {
  createdBy: object;
  createdDateTime: string;
  eTag: string;
  id: string;
  lastModifiedBy: object;
  name: string;
  parentReference: object;
  webUrl: string;
  cTag: string;
  fileSystemInfo: object;
  shared: object;
  size: number;
}

export interface FileDetails extends FileOrFolderDetails {
  //"@microsoft.graph.downloadUrl"
  downloadUrl : string; 
  file: object;
  fileSystemInfo: object;
  media: object;
  photo: object;
  
}



export interface Connection {
    connect(): Promise<OneDriveClient>;
  }

export interface OneDriveClient {
  isConnected(): boolean;
  listContents(): Promise<string>;
  getMySiteUrl(): Promise<string>;
  listSites(): Promise<string>;
  listDrives(siteId:string): Promise<string>;
  listGroups(siteId:string): Promise<string>;
  listDrivesForGroup(groupId:string): Promise<string>;
  listDriveItems(driveId:string): Promise<(FileDetails|FolderDetails)[]>;
  listDriveItemContents(driveId:string, itemId:string): Promise<(FileDetails|FolderDetails)[]>;
  addLocalFileToDrive(driveId:string, folderId:string, filePath:string): Promise<FileDetails>;
  addFileBufferToDrive(driveId:string, folderId:string, filename:string, fileBuffer:Buffer): Promise<FileDetails>;
  createFolderInDrive(driveId: string, folderId: string, folderName: string, conflictBehavior?:ConflictBehavior) : Promise<FolderDetails>;
  copyItemToDrive(driveId:string, folderId:string, itemId:string, conflictBehavior:ConflictBehavior): Promise<string> 


}

export enum ConflictBehavior {
  Replace = 'replace',
  Rename = 'rename',
  Fail = 'fail',
}

export class OneDriveAppFolder implements Connection, OneDriveClient {

  private applicationId: string;
  private applicationSecret: string;
  private tenantId: string;
  private accessToken: string = '';
  private accessTokenExpiration: Date = new Date();


  /**
   * @param applicationId The application ID of the app registration in Azure AD, also known as the client ID. (required)
   * @param applicationSecret The application secret of the app registration in Azure AD, also known as the client secret. (required)
   * @param tenantId The tenant ID of the Azure AD tenant. (required)
   **/
  private constructor(applicationId: string, applicationSecret: string, tenantId: string) { 
    if (applicationId.trim() === '') {
      throw new Error('applicationId is required');
    }
    if (applicationSecret.trim() === '') {
      throw new Error('applicationSecret is required');
    }
    if (tenantId.trim() === '') {
      throw new Error('tenantId is required');
    }

    this.applicationId = applicationId;
    this.applicationSecret = applicationSecret;
    this.tenantId = tenantId;
  }
  public static create(applicationId: string, applicationSecret: string, tenantId: string): Connection {
    return new OneDriveAppFolder(applicationId, applicationSecret, tenantId);
  }

  public async connect(): Promise<OneDriveClient> {
    const tokenResult = await this.requestAccessToken();
    
    this.accessToken = tokenResult.access_token;
    this.accessTokenExpiration = new Date(new Date().getTime() + (tokenResult.expires_in * 1000));
    return this;
  }

  isConnected(): boolean {
    return this.accessToken !== '' && this.accessTokenExpiration > new Date();
  }


  ///site: onedrivepoc

  public async  listContents(): Promise<string> {

    //GET /sites/{siteId}/drives
    const url = `https://graph.microsoft.com/v1.0/sites`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  public async getMySiteUrl(): Promise<string> {
    const url = `https://graph.microsoft.com/v1.0/me?$select=mySite`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  public async listSites(): Promise<string> {
    const url = `https://graph.microsoft.com/v1.0/sites/simnovaoffice.sharepoint.com`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  // list drives in a site:
  public async listDrives(siteId:string): Promise<string> {
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  /** 
   * Requires permission of GroupMember.Read.All on the Entra Application
   */
  public async listGroups(siteId:string): Promise<string> {
    const url = `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')`; 
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  public async listDrivesForGroup(groupId:string): Promise<string> {
    const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/drives`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    }).then(response => response.json())
  }

  public async listDriveItems(driveId:string): Promise<(FileDetails|FolderDetails)[]> {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    })
    .then(response => response.json())
    .then(data => {
      return (data.value as any as [object]).map(item => this.convertToFileOrFolderDetails(item))
    });

  }

  public async listDriveItemContents(driveId:string, itemId:string): Promise<(FileDetails|FolderDetails)[]> {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children`;
    return fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
      },
    })
    .then(response => response.json())
    .then(data => {
      return (data.value as any as [object]).map(item => this.convertToFileOrFolderDetails(item))
    });  }

  public async addLocalFileToDrive(driveId:string, folderId:string, filePath:string): Promise<FileDetails> {
    //get filename from filePath
    const filename = filePath.split('/').pop();
    if (!filename) {
      throw new Error(`Invalid file path: ${filePath}`);
    }
    //use fs to open read stream for file
    const file = fs.readFileSync(filePath);

    return this.addFileBufferToDrive(driveId, folderId, filename, file);
  }

  public async addFileBufferToDrive(driveId:string, folderId:string, filename:string, fileBuffer:Buffer): Promise<FileDetails> {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}:/${filename}:/content`;
    return fetch(url, {
      method: 'PUT',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
        'Content-Type': 'text/plain',
      },
      body: fileBuffer,
    })
    .then(response => response.json())
    .then(data =>  this.convertToFileOrFolderDetails(data) as FileDetails);
  }

  public async copyItemToDrive(driveId:string, folderId:string, itemId:string, conflictBehavior:ConflictBehavior=ConflictBehavior.Replace): Promise<string> {
    console.log('copyItemToDrive', driveId, folderId, itemId, conflictBehavior);
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/copy`;
    const result = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `bearer ${this.accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        parentReference: {
          id: folderId,
        },
        '@microsoft.graph.conflictBehavior': conflictBehavior,
      }),
    })
    .then(response => response.statusText);
    console.log('finishing copyItemToDrive', driveId, folderId, itemId, conflictBehavior);
    return result;
  }

  

  public createFolderInDrive(driveId: string, folderId: string, folderName: string, conflictBehavior:ConflictBehavior=ConflictBehavior.Replace) : Promise<FolderDetails> {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children`;

    return fetch(url, {
      method: 'POST',
      headers: {
      'Authorization': `bearer ${this.accessToken}`,
      'Content-Type': 'application/json',
      },
      body: JSON.stringify({
      name: folderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': conflictBehavior /// ConflictBehavior[conflictBehavior as keyof typeof ConflictBehavior],
      }),
    })
    .then(response => response.json())
    .then(data => {
      return this.convertToFileOrFolderDetails(data) as FolderDetails;
    });
  }


  private convertToFileOrFolderDetails(data: any): FileDetails | FolderDetails {
    //if object has property "folder" then it is a folder, if it has @microsoft.graph.downloadUrl then it is a file, otherwise it is an error
    if (data.hasOwnProperty('folder')) {
      return {
        id: data.id,
        name: data.name,
        webUrl: data.webUrl,
        folder: data.folder,
        createdBy: data.createdBy,
        createdDateTime: data.createdDateTime,
        eTag: data.eTag,
        lastModifiedBy: data.lastModifiedBy,
        parentReference: data.parentReference,
        cTag: data.cTag,
        fileSystemInfo: data.fileSystemInfo,
        shared: data.shared,
        size: data.size
      } as FolderDetails;
    } else if (data.hasOwnProperty('@microsoft.graph.downloadUrl')) {
      return {
        id: data.id,
        name: data.name,
        webUrl: data.webUrl,
        file: data.file,
        createdBy: data.createdBy,
        createdDateTime: data.createdDateTime,
        eTag: data.eTag,
        lastModifiedBy: data.lastModifiedBy,
        parentReference: data.parentReference,
        cTag: data.cTag,
        fileSystemInfo: data.fileSystemInfo,
        shared: data.shared,
        size: data.size,
        downloadUrl: data['@microsoft.graph.downloadUrl'],
        media: data.media,
        photo: data.photo
      } as FileDetails;
    } else {
      throw new Error(`Invalid data : ${JSON.stringify(data)}`);
    }
  }
  


  //https://javascript.plainenglish.io/onedrive-integration-with-react-step-by-step-guide-c068bb8e3fb8


  //https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http
  

  public async requestAccessToken(): Promise<accessTokenResponseType> {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.applicationId,
      client_secret: this.applicationSecret,
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default' //, Sites.Selected', //'api://2a98ceb0-bb4b-4a19-8fab-70464343fd8e/.default', //'https://graph.microsoft.com/.default',
    }).toString();
    const response = await fetch(`https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: body,
      });
      console.log('response', response);
      return response.json();
    }    
}   