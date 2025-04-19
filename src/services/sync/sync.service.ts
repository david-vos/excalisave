import { IDrawing } from "../../interfaces/drawing.interface";
import { XLogger } from "../../lib/logger";
import { browser } from "webextension-polyfill-ts";
import { SyncProvider } from "./sync.interface";

/**
 * Service for managing cloud synchronization of drawings
 */
export class SyncService {
  private static instance: SyncService;
  private provider: SyncProvider | null = null;
  public static readonly SYNC_FOLDER_NAME = "excalidraw-sync";

  private constructor() {}

  /**
   * Get the singleton instance of the SyncService
   */
  public static getInstance(): SyncService {
    if (!SyncService.instance) {
      SyncService.instance = new SyncService();
    }
    return SyncService.instance;
  }

  /**
   * Set the sync provider to use
   */
  public setProvider(provider: SyncProvider): void {
    this.provider = provider;
  }

  /**
   * Initialize the sync service
   */
  public async initialize(): Promise<void> {
    if (!this.provider) {
      XLogger.warn("No sync provider set");
      return;
    }

    try {
      await this.provider.initialize();
      await this.ensureSyncFolderExists();
    } catch (error) {
      XLogger.error("Error initializing sync service", error);
    }
  }

  /**
   * Check if the user is authenticated with the sync provider
   */
  public async isAuthenticated(): Promise<boolean> {
    return (await this.provider?.isAuthenticated()) || false;
  }

  /**
   * Save a drawing to the cloud
   * @param drawing the full class
   */
  public async saveDrawing(drawing: IDrawing): Promise<void> {
    if (!this.provider) return;

    const isAuth = await this.isAuthenticated();
    if (!isAuth) return;

    // Only sync if the drawing is in the sync folder
    const isInSyncFolder = await this.isDrawingInSyncFolder(drawing.id);
    if (!isInSyncFolder) return;

    await this.provider.saveDrawing(drawing);
    XLogger.log("Drawing updated in cloud successfully");
  }

  /**
   * Update an existing drawing in the cloud
   */
  public async updateDrawing(drawing: IDrawing): Promise<void> {
    if (!this.provider) return;

    const isAuth = await this.isAuthenticated();
    if (!isAuth) return;

    // Only sync if the drawing is in the sync folder
    const isInSyncFolder = await this.isDrawingInSyncFolder(drawing.id);
    if (!isInSyncFolder) return;

    await this.provider.updateDrawing(drawing);
    XLogger.log("Drawing updated in cloud successfully");
  }

  /**
   * Check if a drawing is in the sync folder
   */
  private async isDrawingInSyncFolder(drawingId: string): Promise<boolean> {
    try {
      XLogger.info(`Checking if drawing is in sync folder: ${drawingId}`);
      const folders = await browser.storage.local.get("folders");
      const syncFolder = folders.folders?.find(
        (f: any) => f.name === SyncService.SYNC_FOLDER_NAME
      );

      const isInSyncFolder =
        syncFolder?.drawingIds?.includes(drawingId) || false;
      XLogger.info(
        `Drawing ${drawingId} is ${isInSyncFolder ? "" : "not"} in sync folder`
      );
      return isInSyncFolder;
    } catch (error) {
      XLogger.error("Error checking if drawing is in sync folder", error);
      return false;
    }
  }

  /**
   * Ensure the sync folder exists in local storage
   */
  private async ensureSyncFolderExists(): Promise<void> {
    try {
      const folders = await browser.storage.local.get("folders");
      const syncFolder = folders.folders?.find(
        (f: any) => f.name === SyncService.SYNC_FOLDER_NAME
      );

      if (!syncFolder) {
        const newFolder = {
          id: `folder:${Math.random().toString(36).substr(2, 9)}`,
          name: SyncService.SYNC_FOLDER_NAME,
          drawingIds: [] as string[],
        };

        const newFolders = [...(folders.folders || []), newFolder];
        await browser.storage.local.set({ folders: newFolders });
      }
    } catch (error) {
      XLogger.error("Error ensuring sync folder exists", error);
    }
  }

  /**
   * Delete a drawing from the cloud
   */
  public async deleteDrawing(drawingName: string): Promise<void> {
    XLogger.info(`Attempting to delete drawing from cloud: ${drawingName}`);
    if (!this.provider) {
      XLogger.info("No provider set, skipping cloud delete");
      return;
    }

    const isAuth = await this.isAuthenticated();
    if (!isAuth) {
      XLogger.info("User not authenticated, skipping cloud delete");
      return;
    }

    try {
      // Only sync if the drawing is in the sync folder
      const isInSyncFolder = await this.isDrawingInSyncFolder(drawingName);
      if (!isInSyncFolder) {
        XLogger.info(
          `Drawing ${drawingName} is not in sync folder, skipping cloud delete`
        );
        return;
      }

      XLogger.info(`Deleting drawing from cloud: ${drawingName}`);
      await this.provider.deleteDrawing(drawingName);
      XLogger.info(`Drawing deleted from cloud successfully: ${drawingName}`);
    } catch (error) {
      XLogger.error("Error deleting drawing from cloud", error);
    }
  }

  /**
   * Sync files between local and cloud
   */
  public async syncFiles(): Promise<void> {
    if (!this.provider) return;

    const isAuth = await this.isAuthenticated();
    if (!isAuth) return;

    try {
      await this.provider.syncFiles();
    } catch (error) {
      XLogger.error("Error syncing files", error);
    }
  }
}
