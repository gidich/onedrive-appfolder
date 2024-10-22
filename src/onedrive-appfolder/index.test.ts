import { expect, test, describe, beforeEach } from "@jest/globals";
import { OneDriveAppFolder, OneDriveClient } from "./index";

const applicationId = process.env.ONEDRIVE_APP_ID??"";
const applicationSecret = process.env.ONEDRIVE_APP_SECRET??"";
const tenantId = process.env.ONEDRIVE_TENANT_ID??"";
const driveId = process.env.ONEDRIVE_DRIVE_ID??"";
const rootFolderId = process.env.ONEDRIVE_ROOT_FOLDER_ID??"";



describe("When using the onedrive-appfolder module", () => {
    let client: OneDriveClient;
    beforeEach(async () => {
        client = await OneDriveAppFolder.create(applicationId, applicationSecret, tenantId).connect();
    });


    test.skip('oneDriveAppFolder should raise an error if constructed with an empty applicationId ', () => {
        // Arrange
        // Act
        // Assert
        expect(() => OneDriveAppFolder.create('', applicationSecret, tenantId)).toThrowError('applicationId is required');
    });
    test.skip('oneDriveAppFolder should raise an error if constructed with an empty applicationSecret ', () => {
        // Arrange
        // Act
        // Assert
        expect(() => OneDriveAppFolder.create(applicationId, '', tenantId)).toThrowError('applicationSecret is required');
    });
    test('oneDriveAppFolder should raise an error if constructed with an empty tenantId ', () => {
        // Arrange
        // Act
        // Assert
        expect(() => OneDriveAppFolder.create(applicationId, applicationSecret, '')).toThrowError('tenantId is required');
    });

    test.skip('should be connected after instantiation', async () => {
        // Arrange
        const appFolder = OneDriveAppFolder.create(applicationId, applicationSecret, tenantId);
        
        // Act
        const client = await appFolder.connect();
        // Assert
        expect(client.isConnected()).toBe(true);
    });

    test.skip('should be able to get my site url', async () => {
        // Arrange
        // Act
        const result = await client.getMySiteUrl();
        console.log('get my site url result', result);  
        // Assert
        expect(result).toBeDefined();
    });



    test.skip('should be able to list sites', async () => {
        // Arrange
        // Act
        const result = await client.listSites();
        console.log('list sites result', result);  
        // Assert
        expect(result).toBeDefined();
    });

    test.skip('should be able to list drives for a site', async () => {

        // Arrange
        const siteId = 'simnovaoffice.sharepoint.com,ebeec291-776d-474e-9a2f-8c88749ed244,a0496189-e5d8-458b-ac25-24b312a0b86b';
        // Act
        const result = await client.listDrives(siteId);
        console.log('list drives result:', result);  
        // Assert
        expect(result).toBeDefined();
    });

    test.skip('should be able to list groups for a site', async () => {
        // Arrange
        const siteId = 'simnovaoffice.sharepoint.com,ebeec291-776d-474e-9a2f-8c88749ed244,a0496189-e5d8-458b-ac25-24b312a0b86b';
        // Act
        const result = await client.listGroups(siteId);
        console.log('list groups result', result);  
        // Assert
        expect(result).toBeDefined();
    });

    test.skip('should be able to list drives for a group', async () => { 
        // Arrange
        const groupId = 'ad8a6a08-f6ee-483e-8a98-0ff628242c81'; //oneDrivePoc group
        // Act
        const result = await client.listDrivesForGroup(groupId);
        console.log('list drives for group result', result);  
        // Assert
        expect(result).toBeDefined();
    });

    test.skip('should be able to list items in a drive', async () => {
        // Arrange
        // Act
        const result = await client.listDriveItems(driveId);
        console.log('list drive items result', result);
        // Assert
        expect(result).toBeDefined();
    });


    test.skip('should be able to add a local file to the drive', async () => {   
        // Arrange
        const folderId = rootFolderId;
        const localFilePath = 'src/onedrive-appfolder/index.ts';
        // Act
        const result = await client.addLocalFileToDrive(driveId, folderId, localFilePath);
        console.log('upload file result', result);
        // Assert
        expect(result).toBeDefined();
    });


    test.skip('should be able to list items in a folder', async () => {
        // Arrange
        const folderId = rootFolderId;
        // Act
        const result = await client.listDriveItemContents(driveId, folderId);
        console.log('list folder items result', result);
        // Assert
        expect(result).toBeDefined();
    });


    


    test.skip('should be able to list contents of app folder', async () => {
        // Arrange
        // Act
        const result = await client.listContents();
        console.log('list contents result', result);  
        // Assert
        expect(result).toBeDefined();
    });

    test.skip('should be able to create an empty folder in the drive', async () => {
        // Arrange
        const folderId = rootFolderId;
        const folderName = 'user@host.com';
        // Act
        const result = await client.createFolderInDrive(driveId, folderId, folderName);
        console.log('create folder result', result);
        // Assert
        expect(result).toBeDefined();
    });
    test.skip('should be able to create a set of folders when creating a root folder', async () => { 
        // Arrange
        const folderId = rootFolderId;
        //create a unique folder name for each run, should be in the format of an email address:
        const folderName = '' + new Date().toISOString().replace(/:/g, '-').replace(/\./g, '-').replace(/T/g, '-').replace(/Z/g, '').replace(/-/g, '') + '@host.com';
        // Act
        // create root folder
        const rootFolder = await client.createFolderInDrive(driveId, folderId, folderName);
        // create 3 supporting folders ('shared-applicant-editable', 'shared-applicant-readonly', 'private')
        const sharedApplicantEditableFolder = await client.createFolderInDrive(driveId, rootFolder.id, 'shared-applicant-editable');
        const sharedApplicantReadonlyFolder = await client.createFolderInDrive(driveId, rootFolder.id, 'shared-applicant-readonly');
        const privateFolder = await client.createFolderInDrive(driveId, rootFolder.id, 'private');
        // Assert
        expect(rootFolder).toBeDefined();
        expect(sharedApplicantEditableFolder).toBeDefined();
        expect(sharedApplicantReadonlyFolder).toBeDefined();
        expect(privateFolder).toBeDefined();
    });



});
