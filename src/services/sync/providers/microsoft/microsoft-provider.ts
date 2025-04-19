import { IDrawing } from "../../../../interfaces/drawing.interface";
import { XLogger } from "../../../../lib/logger";
import { browser } from "webextension-polyfill-ts";
import { SyncProvider } from "../../sync.interface";
import { MicrosoftAuthService } from "./microsoft-auth.service";
import { MicrosoftFetchService } from "./microsoft-fetch.service";

/**
 * Microsoft OneDrive implementation of the SyncProvider interface
 */
export class MicrosoftProvider implements SyncProvider {
  private static instance: MicrosoftProvider;
  public static readonly SYNC_FOLDER_NAME = "excalidraw-sync";
  private authService: MicrosoftAuthService;

  private constructor() {
    this.authService = MicrosoftAuthService.getInstance();
  }

  /**
   * Get the singleton instance of the MicrosoftProvider
   */
  public static getInstance(): MicrosoftProvider {
    if (!MicrosoftProvider.instance) {
      MicrosoftProvider.instance = new MicrosoftProvider();
    }
    return MicrosoftProvider.instance;
  }

  /**
   * Initialize the Microsoft provider
   */
  public async initialize(): Promise<void> {
    try {
      const isAuth = await this.authService.isAuthenticated();
      if (!isAuth) return;

      await this.syncFiles();
    } catch (error) {
      XLogger.error(
        `Error initializing Microsoft provider: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  /**
   * Check if the user is authenticated with Microsoft
   */
  public async isAuthenticated(): Promise<boolean> {
    return await this.authService.isAuthenticated();
  }

  /**
   * Saves or updates a drawing to Microsoft OneDrive
   * @param drawing The drawing to save or update
   * @param isUpdate Whether this is an update operation (for logging purposes)
   */
  private async saveOrUpdateDrawing(
    drawing: IDrawing,
    isUpdate: boolean = false
  ): Promise<void> {
    try {
      await this.ensureSyncFolderExists();
      const accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        throw new Error("Access token is required");
      }

      await MicrosoftFetchService.saveFile(
        MicrosoftProvider.SYNC_FOLDER_NAME,
        `${drawing.name}.json`,
        drawing,
        accessToken
      );
    } catch (error) {
      const operation = isUpdate ? "updating" : "saving";
      XLogger.error(
        `Error ${operation} drawing to Microsoft OneDrive: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }

  /**
   * Save a drawing to Microsoft OneDrive
   */
  public async saveDrawing(drawing: IDrawing): Promise<void> {
    await this.saveOrUpdateDrawing(drawing, false);
  }

  /**
   * Update an existing drawing in Microsoft OneDrive
   */
  public async updateDrawing(drawing: IDrawing): Promise<void> {
    await this.saveOrUpdateDrawing(drawing, true);
  }

  /**
   * Delete a drawing from Microsoft OneDrive
   */
  public async deleteDrawing(drawingName: string): Promise<void> {
    try {
      XLogger.info(
        `Attempting to delete drawing from OneDrive: ${drawingName}`
      );
      await this.ensureSyncFolderExists();
      const accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        throw new Error("Access token is required");
      }

      XLogger.info(`Sending delete request to OneDrive for: ${drawingName}`);
      const response = await MicrosoftFetchService.deleteFile(
        MicrosoftProvider.SYNC_FOLDER_NAME,
        `${drawingName}.json`,
        accessToken
      );

      XLogger.info(
        `Delete response status: ${response.status} ${response.statusText}`
      );
      XLogger.info(
        `Drawing deleted from OneDrive successfully: ${drawingName}`
      );
    } catch (error) {
      XLogger.error(
        `Error deleting drawing from Microsoft OneDrive: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }

  /**
   * Sync files between local storage and Microsoft OneDrive
   */
  public async syncFiles(): Promise<void> {
    try {
      XLogger.info("Starting sync from OneDrive to local storage");
      const oneDriveDrawings = await this.getOneDriveFiles();
      XLogger.info(
        `Retrieved ${oneDriveDrawings.length} drawings from OneDrive`
      );

      const folders = await browser.storage.local.get("folders");
      const oneDriveFolder = folders.folders?.find(
        (f: any) => f.name === MicrosoftProvider.SYNC_FOLDER_NAME
      );

      if (!oneDriveFolder) {
        XLogger.error("OneDrive folder not found in local storage");
        return;
      }

      XLogger.info(
        `Found OneDrive folder in local storage with ${
          oneDriveFolder.drawingIds?.length || 0
        } drawings`
      );

      // Update local storage with OneDrive files
      for (const drawing of oneDriveDrawings) {
        XLogger.info(
          `Processing OneDrive drawing: ${drawing.name} (ID: ${drawing.id})`
        );
        const localDrawing = (await browser.storage.local.get(drawing.id))[
          drawing.id
        ];

        // Check for conflicts
        if (
          localDrawing &&
          localDrawing.data.versionFiles !== drawing.data.versionFiles
        ) {
          XLogger.info(`Conflict detected for drawing: ${drawing.name}`);
          // Emit conflict event for UI to handle
          await browser.runtime.sendMessage({
            type: "MICROSOFT_SYNC_CONFLICT",
            payload: {
              drawingId: drawing.id,
              localDrawing,
              oneDriveDrawing: drawing,
            },
          });
          continue;
        }

        XLogger.info(
          `Saving OneDrive drawing to local storage: ${drawing.name}`
        );
        await browser.storage.local.set({
          [drawing.id]: drawing,
        });

        // Add to OneDrive folder if not already there
        if (!oneDriveFolder.drawingIds.includes(drawing.id)) {
          XLogger.info(`Adding drawing to OneDrive folder: ${drawing.name}`);
          oneDriveFolder.drawingIds.push(drawing.id);
        }
      }

      // Save updated folder
      XLogger.info("Saving updated OneDrive folder to local storage");
      await browser.storage.local.set({ folders: folders.folders });
      XLogger.info("Sync from OneDrive to local storage completed");
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      XLogger.error(`Error syncing OneDrive files: ${errorMessage}`);
      throw error;
    }
  }

  /**
   * Ensure the sync folder exists in Microsoft OneDrive
   */
  private async ensureSyncFolderExists(): Promise<string> {
    try {
      const isAuth = await this.authService.isAuthenticated();
      if (!isAuth) {
        throw new Error("User is not authenticated");
      }

      const accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        throw new Error("Access token is required");
      }

      XLogger.info(
        `Checking if folder exists: ${MicrosoftProvider.SYNC_FOLDER_NAME}`
      );
      // Check if folder exists
      try {
        const response = await MicrosoftFetchService.checkFolderExists(
          MicrosoftProvider.SYNC_FOLDER_NAME,
          accessToken
        );
        const folderData = await response.json();
        XLogger.info(`Folder exists with ID: ${folderData.id}`);
        return folderData.id;
      } catch (error) {
        // Folder doesn't exist, create it
        XLogger.info(
          `Folder doesn't exist, creating it: ${MicrosoftProvider.SYNC_FOLDER_NAME}`
        );
        const createResponse = await MicrosoftFetchService.createFolder(
          MicrosoftProvider.SYNC_FOLDER_NAME,
          accessToken
        );
        const folderData = await createResponse.json();
        XLogger.info(`Created folder with ID: ${folderData.id}`);
        return folderData.id;
      }
    } catch (error) {
      XLogger.error(
        `Error ensuring sync folder exists: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }

  /**
   * Get all drawings from Microsoft OneDrive
   */
  private async getOneDriveFiles(): Promise<IDrawing[]> {
    try {
      const isAuth = await this.authService.isAuthenticated();
      if (!isAuth) {
        throw new Error("User is not authenticated");
      }

      const accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        throw new Error("Access token is required");
      }

      const folderId = await this.ensureSyncFolderExists();
      XLogger.info(`Retrieved folder ID: ${folderId}`);

      const response = await MicrosoftFetchService.getFolderItems(
        folderId,
        accessToken
      );
      const data = await response.json();
      XLogger.info(`Retrieved ${data.value.length} items from OneDrive folder`);

      const files = data.value.filter((file: any) =>
        file.name.endsWith(".json")
      );
      XLogger.info(`Found ${files.length} JSON files in OneDrive folder`);

      const drawings: IDrawing[] = [];
      for (const file of files) {
        try {
          XLogger.info(
            `Getting content for file: ${file.name} (ID: ${file.id})`
          );
          // getFileContent already returns the parsed JSON
          const content = await MicrosoftFetchService.getFileContent(
            file.id,
            accessToken
          );
          XLogger.info(`Successfully retrieved content for file: ${file.name}`);
          drawings.push(content);
        } catch (error) {
          XLogger.error(
            `Error getting content for file ${file.name}: ${
              error instanceof Error ? error.message : String(error)
            }`
          );
        }
      }

      XLogger.info(
        `Successfully retrieved ${drawings.length} drawings from OneDrive`
      );
      return drawings;
    } catch (error) {
      XLogger.error(
        `Error getting OneDrive files: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }
}
