import { OneDriveClient, FolderDetails, FileDetails, ConflictBehavior } from "../onedrive-appfolder";


export interface SharedFileDetails {
    name: string;
    id: string;
    createdDateTime: string;
    size: number;
    mimeType: string;
}

export interface ApplicantFolderContents {
    sharedApplicantEditable: SharedFileDetails[],
    sharedApplicantReadonly: SharedFileDetails[],
}

export class SampleApp {
    private readonly templateFolderName = '_template';
    private readonly parentFolderName = 'applicant-root';
    private readonly sharedApplicantEditableFolderName = 'shared-applicant-editable';
    private readonly sharedApplicantReadonlyFolderName = 'shared-applicant-readonly';
    private readonly client: OneDriveClient;
    private readonly rootFolderId: string;
    private readonly driveId: string;
    
    public async uploadFileToSharedApplicantEditableFolder(applicantEmail: string, fileName: string, fileBuffer: Buffer): Promise<SharedFileDetails> {
        const applicantFolderName = applicantEmail;
        if(!this.client.isConnected()){
            throw new Error('Not connected to OneDrive');
        }
        const parentFolder = await this.createFolderIfNotExists(this.driveId, this.rootFolderId, this.parentFolderName);
        const applicantFolder = await this.createFolderIfNotExists(this.driveId, parentFolder.id, applicantFolderName);
        const sharedApplicantEditableFolder = await this.createFolderIfNotExists(this.driveId, applicantFolder.id, this.sharedApplicantEditableFolderName);
        const fileDetails = await this.client.addFileBufferToDrive(this.driveId, sharedApplicantEditableFolder.id, fileName, fileBuffer);
        return this.convertToSharedFileDetails(fileDetails);
    }


    constructor(oneDriveAppFolder: OneDriveClient, rootFolderId: string, driveId: string) {
        this.client = oneDriveAppFolder;
        this.rootFolderId = rootFolderId;
        this.driveId = driveId;
    }

    /**
     * 
     * @param applicantEmail 
     * 
     * This method should create a folder structure for a new applicant. The folder structure should be as follows:
     *  - A folder named after the applicant's email address
     *  - 3 subfolders within the applicant's folder: 'shared-applicant-editable', 'shared-applicant-readonly', 'private'
     * @returns returns an object containing the listing of files in each of the shared folders
     */
    async initializeFolderStructureForApplicant(applicantEmail: string) : Promise<ApplicantFolderContents> {
        const applicantFolderName = applicantEmail;
        const privateFolderName = 'private';


        if(!this.client.isConnected()){
            throw new Error('Not connected to OneDrive');
        }

        const parentFolder = await this.createFolderIfNotExists(this.driveId, this.rootFolderId, this.parentFolderName);
        const applicantFolder = await this.createFolderIfNotExists(this.driveId, parentFolder.id, applicantFolderName);

        const subfolders = await this.createSubfoldersIfNeeded(this.driveId, applicantFolder.id, [this.sharedApplicantEditableFolderName, this.sharedApplicantReadonlyFolderName, privateFolderName]);
        const sharedApplicantEditableFolder = subfolders.find((folder) => folder.name === this.sharedApplicantEditableFolderName);
        const sharedApplicantReadonlyFolder = subfolders.find((folder) => folder.name === this.sharedApplicantReadonlyFolderName);

        if(!sharedApplicantEditableFolder || !sharedApplicantReadonlyFolder){
            console.log('subfolders', subfolders);
            throw new Error('Could not create subfolders');
        }
       
        return {
            sharedApplicantEditable: (await this.listFilesInFolder(this.driveId, sharedApplicantEditableFolder.id)).map((file) => this.convertToSharedFileDetails(file)),
            sharedApplicantReadonly: (await this.listFilesInFolder(this.driveId, sharedApplicantReadonlyFolder.id)).map((file) => this.convertToSharedFileDetails(file)),
        };
    }

    private async populateSubfolderWithFilesFromTemplate(subfolder: FolderDetails) : Promise<void> {
        const templateFolder = await this.createFolderIfNotExists(this.driveId, this.rootFolderId, this.templateFolderName);
        const templateSubfolder = await this.createFolderIfNotExists(this.driveId, templateFolder.id, subfolder.name);

        const templateFiles = await this.listFilesInFolder(this.driveId, templateSubfolder.id);
       // console.log(`template files for folder: ${subfolder.name} `, templateFiles);
        const result = await Promise.all(templateFiles.map(async (file) => {
            try{
                const result = await this.client.copyItemToDrive(this.driveId, subfolder.id, file.id, ConflictBehavior.Fail)
                console.log(`copied file ${file.name} to folder ${subfolder.name}`, result);
            }
            catch(e){
                console.log(`Error copying file ${file.name} to folder ${subfolder.name}`, e);
            }
        }   
        ));
        return;
    }
    private async createSubfoldersIfNeeded(driveId: string, parentFolderId: string, subfolderNames: string[]) : Promise<FolderDetails[]> {
        const subfolders = await this.listFoldersInFolder(driveId, parentFolderId);
        const matchedSubfolders = subfolders.filter((folder) => subfolderNames.includes(folder.name));
        const unmatchedSubfolderNames = subfolderNames.filter((name) => !matchedSubfolders.find((folder) => folder.name === name));
        console.log('unmatched subfolders', unmatchedSubfolderNames);
        console.log('matched subfolders', matchedSubfolders);
        const createdSubFolders = await Promise.all(unmatchedSubfolderNames.map(async (name) => {
            const newFolder = await this.client.createFolderInDrive(driveId, parentFolderId, name, ConflictBehavior.Fail);
            await this.populateSubfolderWithFilesFromTemplate(newFolder);
            return newFolder;
        }));
        return matchedSubfolders.concat( createdSubFolders);
    }

    isFolderDetails(item: any): item is FolderDetails {
        return(item as FolderDetails).folder !== undefined;
    }

    private async listFilesInFolder(driveId: string, folderId: string) : Promise<FileDetails[]> {
        const contents = await this.client.listDriveItemContents(driveId, folderId);
        return contents.filter((item) => !this.isFolderDetails(item)) as FileDetails[];
    }
    private async listFoldersInFolder(driveId: string, folderId: string) : Promise<FolderDetails[]> {
        const contents = await this.client.listDriveItemContents(driveId, folderId);
        return contents.filter((item) => this.isFolderDetails(item)) as FolderDetails[];
    }

    private convertToSharedFileDetails(file: FileDetails) : SharedFileDetails {
        const {mimeType} = file.file as any;
        return {
            name: file.name,
            id: file.id,
            createdDateTime: file.createdDateTime,
            size: file.size,
            mimeType: mimeType,
        };
    };

    private async createFolderIfNotExists(driveId: string, parentFolderId: string, folderName: string) : Promise<FolderDetails> {
        const contents = await this.client.listDriveItemContents(driveId, parentFolderId);
        const folder = contents.find((item) => item.name === folderName);
        if(folder){
            return folder as FolderDetails;
        }
        return this.client.createFolderInDrive(driveId, parentFolderId, folderName, ConflictBehavior.Fail);
    }

}