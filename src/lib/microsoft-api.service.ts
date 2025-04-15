import { IDrawing } from "../interfaces/drawing.interface";
import { XLogger } from "./logger";
import { browser } from "webextension-polyfill-ts";
import { MicrosoftFetchService } from "./microsoft-fetch.service";
import { MicrosoftAuthService } from "./microsoft-auth.service";

export class MicrosoftApiService {
  private static instance: MicrosoftApiService;
  public static readonly SYNC_FOLDER_NAME = "excalidraw-microsoft-sync";

  private constructor() {}

  public static getInstance(): MicrosoftApiService {
    if (!MicrosoftApiService.instance) {
      MicrosoftApiService.instance = new MicrosoftApiService();
    }
    return MicrosoftApiService.instance;
  }

  private static async initializePeriodicSync() {
    // Sync every 5 minutes
    setInterval(
      async () => {
        try {
          await this.syncOneDriveFiles();
        } catch (error) {
          XLogger.error("Error during periodic sync", error);
        }
      },
      5 * 60 * 1000
    );
  }

  public static async initialize(): Promise<void> {
    await MicrosoftAuthService.ensureInitialized();

    if (MicrosoftAuthService.isAuthenticated()) {
      await this.initializePeriodicSync();
    }
  }

  private static async ensureSyncFolderExists(): Promise<string> {
    try {
      if (!MicrosoftAuthService.accessToken) {
        throw new Error("Access token is required");
      }

      // Check if folder exists
      try {
        const response = await MicrosoftFetchService.checkFolderExists(
          MicrosoftApiService.SYNC_FOLDER_NAME,
          MicrosoftAuthService.accessToken
        );
        const folderData = await response.json();
        return folderData.id;
      } catch (error) {
        // Folder doesn't exist, create it
        const createResponse = await MicrosoftFetchService.createFolder(
          MicrosoftApiService.SYNC_FOLDER_NAME,
          MicrosoftAuthService.accessToken
        );
        const folderData = await createResponse.json();
        return folderData.id;
      }
    } catch (error) {
      XLogger.error("Error ensuring sync folder exists", error);
      throw error;
    }
  }

  private static async getOneDriveFiles(): Promise<IDrawing[]> {
    try {
      if (!MicrosoftAuthService.accessToken) {
        throw new Error("Access token is required");
      }

      const folderId = await this.ensureSyncFolderExists();
      const response = await MicrosoftFetchService.getFolderItems(
        folderId,
        MicrosoftAuthService.accessToken
      );

      const data = await response.json();
      const files = data.value.filter((file: any) =>
        file.name.endsWith(".json")
      );

      const drawings: IDrawing[] = [];
      for (const file of files) {
        try {
          const drawing = await MicrosoftFetchService.getFileContent(
            file.id,
            MicrosoftAuthService.accessToken
          );
          drawings.push(drawing);
        } catch (error) {
          XLogger.error(`Error getting content for file ${file.name}`, error);
        }
      }

      return drawings;
    } catch (error) {
      XLogger.error("Error getting OneDrive files", error);
      throw error;
    }
  }

  private static async isDrawingInOneDriveFolder(
    drawingId: string
  ): Promise<boolean> {
    try {
      const folders = await browser.storage.local.get("folders");
      const oneDriveFolder = folders.folders?.find(
        (f: any) => f.name === this.SYNC_FOLDER_NAME
      );

      return oneDriveFolder && oneDriveFolder.drawingIds.includes(drawingId);
    } catch (error) {
      XLogger.error("Error checking if drawing is in OneDrive folder", error);
      return false;
    }
  }

  private static async performSyncAfterAction(): Promise<void> {
    try {
      await this.syncOneDriveFiles();
    } catch (error) {
      XLogger.error("Error syncing after action", error);
    }
  }

  public static async syncOneDriveFiles(): Promise<void> {
    await MicrosoftAuthService.ensureInitialized();

    try {
      const oneDriveDrawings = await this.getOneDriveFiles();
      const folders = await browser.storage.local.get("folders");
      const oneDriveFolder = folders.folders?.find(
        (f: any) => f.name === this.SYNC_FOLDER_NAME
      );

      if (!oneDriveFolder) {
        XLogger.error("OneDrive folder not found in local storage");
        return;
      }

      // Update local storage with OneDrive files
      for (const drawing of oneDriveDrawings) {
        await browser.storage.local.set({
          [drawing.id]: drawing,
        });

        // Add to OneDrive folder if not already there
        if (!oneDriveFolder.drawingIds.includes(drawing.id)) {
          oneDriveFolder.drawingIds.push(drawing.id);
        }
      }

      // Save updated folder
      await browser.storage.local.set({ folders: folders.folders });
    } catch (error) {
      XLogger.error("Error syncing OneDrive files", error);
      throw error;
    }
  }

  static async saveDrawing(drawing: IDrawing): Promise<void> {
    await MicrosoftAuthService.ensureInitialized();

    try {
      // Only sync if the drawing is in the correct folder
      const isInOneDriveFolder = await this.isDrawingInOneDriveFolder(
        drawing.id
      );
      if (!isInOneDriveFolder) {
        return; // Skip sync if not in OneDrive folder
      }

      await this.ensureSyncFolderExists();

      await MicrosoftFetchService.saveFile(
        MicrosoftApiService.SYNC_FOLDER_NAME,
        `${drawing.name}.json`,
        drawing,
        MicrosoftAuthService.accessToken!
      );

      XLogger.log("Drawing saved to Microsoft OneDrive successfully");

      // Sync OneDrive files after save
      await this.performSyncAfterAction();
    } catch (error) {
      XLogger.error("Error saving drawing to Microsoft OneDrive", error);
      throw error;
    }
  }

  static async updateDrawing(drawing: IDrawing): Promise<void> {
    await MicrosoftAuthService.ensureInitialized();

    try {
      // Only sync if the drawing is in the OneDrive folder
      const isInOneDriveFolder = await this.isDrawingInOneDriveFolder(
        drawing.id
      );
      if (!isInOneDriveFolder) {
        return; // Skip sync if not in OneDrive folder
      }

      await this.ensureSyncFolderExists();

      await MicrosoftFetchService.saveFile(
        MicrosoftApiService.SYNC_FOLDER_NAME,
        `${drawing.name}.json`,
        drawing,
        MicrosoftAuthService.accessToken!
      );

      XLogger.log("Drawing updated in Microsoft OneDrive successfully");

      // Sync OneDrive files after update
      await this.performSyncAfterAction();
    } catch (error) {
      XLogger.error("Error updating drawing in Microsoft OneDrive", error);
      throw error;
    }
  }

  static async deleteDrawing(drawingName: string): Promise<void> {
    await MicrosoftAuthService.ensureInitialized();

    try {
      // Only sync if the drawing is in the OneDrive folder
      const isInOneDriveFolder =
        await this.isDrawingInOneDriveFolder(drawingName);
      if (!isInOneDriveFolder) {
        return; // Skip sync if not in OneDrive folder
      }

      await this.ensureSyncFolderExists();

      await MicrosoftFetchService.deleteFile(
        MicrosoftApiService.SYNC_FOLDER_NAME,
        `${drawingName}.json`,
        MicrosoftAuthService.accessToken!
      );

      XLogger.log("Drawing deleted from Microsoft OneDrive successfully");

      // Sync OneDrive files after delete
      await this.performSyncAfterAction();
    } catch (error) {
      XLogger.error("Error deleting drawing from Microsoft OneDrive", error);
      throw error;
    }
  }
}
