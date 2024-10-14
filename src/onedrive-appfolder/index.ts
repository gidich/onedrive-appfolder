export class OneDriveAppFolder {
  constructor(applicationId: string, applicationSecret: string) {
    console.log('OneDriveAppFolder constructor');
  }
  isConnected(): boolean {
    return true;
  }
}   