import { expect, test, describe, beforeEach } from "@jest/globals";
import { OneDriveAppFolder } from "./index";

const applicationId = process.env.ONEDRIVE_APP_ID??"";
const applicationSecret = process.env.ONEDRIVE_APP_SECRET??"";


describe("When using the onedrive-appfolder module", () => {
    let oneDriveAppFolder: any;
    beforeEach(() => {
        oneDriveAppFolder = new OneDriveAppFolder(applicationId, applicationSecret);
    });
    test('should be connected after instantiation', () => {

        expect(oneDriveAppFolder.isConnected()).toBe(true);
    });
});
