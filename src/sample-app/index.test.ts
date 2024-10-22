import { expect, test, describe, beforeEach } from "@jest/globals";
import { SampleApp } from "./index";
import { on } from "events";
import { OneDriveAppFolder, OneDriveClient } from "../onedrive-appfolder";
import path from "path";
import fs from "fs";

const applicationId = process.env.ONEDRIVE_APP_ID??"";
const applicationSecret = process.env.ONEDRIVE_APP_SECRET??"";
const tenantId = process.env.ONEDRIVE_TENANT_ID??"";
const driveId = process.env.ONEDRIVE_DRIVE_ID??"";
const rootFolderId = process.env.ONEDRIVE_ROOT_FOLDER_ID??"";

describe("When using the onedrive-appfolder module", () => {





    let client: OneDriveClient
    let sampleApp: SampleApp;

    beforeEach(async () => {
        client = await OneDriveAppFolder.create(applicationId, applicationSecret, tenantId).connect();
        sampleApp = new SampleApp(client, rootFolderId, driveId);
    });

    test('initialize folder structure for a new user', async () => {
        // Arrange
        const applicantEmail =   '' + new Date().toISOString().replace(/:/g, '-').replace(/\./g, '-').replace(/T/g, '-').replace(/Z/g, '').replace(/-/g, '') + '@host.com';
        
        // Act
        const result = await sampleApp.initializeFolderStructureForApplicant(applicantEmail);
        // Assert
    
        
        //confirm that a folder is created for the applicant matching their email address and that 3 subfolders are created: ('shared-applicant-editable', 'shared-applicant-readonly', 'private'), use result confirm
        console.log('initialize folder structure result', result);
        expect(result).toBeDefined();
    }, 10000);

    test('return folder structure for an existing user', async () => {
        // Arrange
        const applicantEmail =   'jane-doe@host.com';
        
        // Act
        const startTime = performance.now();
        const result = await sampleApp.initializeFolderStructureForApplicant(applicantEmail);
        const endTime = performance.now();
        console.log(`Time taken to get folder structure for existing user: ${(endTime - startTime)/1000}s`);
        // Assert
    
        
        //confirm that a folder is created for the applicant matching their email address and that 3 subfolders are created: ('shared-applicant-editable', 'shared-applicant-readonly', 'private'), use result confirm
        console.log('initialize folder structure result', result);
        expect(result).toBeDefined();        
    });

    test('upload file to shared-applicant-editable folder', async () => {
        // Arrange
        const applicantEmail =   'jane-doe@host.com';
        const filePath = 'src/sample-app/index.test.ts';
        const fileBuffer = fs.readFileSync(filePath);
        const fileName = path.basename(filePath);

        // Act
        const startTime = performance.now();
        const result = await sampleApp.uploadFileToSharedApplicantEditableFolder(applicantEmail, fileName, fileBuffer);
        const endTime = performance.now();
        console.log(`Time taken to upload file to shared-applicant-editable folder: ${(endTime - startTime)/1000}s`);

        // Assert
        console.log('upload file result', result);
        expect(result).toBeDefined();
    });

});