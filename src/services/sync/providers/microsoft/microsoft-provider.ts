import { IDrawing } from "../../../../interfaces/drawing.interface";
import { XLogger } from "../../../../lib/logger";
import { browser } from "webextension-polyfill-ts";
import { SyncProvider } from "../../sync.interface";
import { MicrosoftAuthService } from "./microsoft-auth.service";
import { MicrosoftFetchService } from "./microsoft-fetch.service";
import { DeltaSyncService } from "../../delta-sync.service";
import { SyncStatus } from "../../../../interfaces/sync-status.interface";

/**
 * Microsoft OneDrive implementation of the SyncProvider interface
 */
export class MicrosoftProvider implements SyncProvider {
  private static instance: MicrosoftProvider;
  public static readonly SYNC_FOLDER_NAME = "excalidraw-sync";
  private authService: MicrosoftAuthService;
  private deltaSyncService: DeltaSyncService;

  private constructor() {
    this.authService = MicrosoftAuthService.getInstance();
    this.deltaSyncService = DeltaSyncService.getInstance();
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
      // Update sync status to syncing
      this.deltaSyncService.updateSyncStatus(drawing.id, {
        status: SyncStatus.SYNCING,
        version: drawing.data.versionFiles,
        progress: 0,
      });

      await this.ensureSyncFolderExists();
      const accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        throw new Error("Access token is required");
      }

      // Get the current version from OneDrive if it exists
      let currentVersion: IDrawing | null = null;
      try {
        currentVersion = await MicrosoftFetchService.getFileContentByPath(
          `${MicrosoftProvider.SYNC_FOLDER_NAME}/${drawing.name}.json`,
          accessToken
        );
      } catch (error) {
        // File doesn't exist yet, which is fine
      }

      if (currentVersion) {
        // Calculate and apply delta
        try {
          const delta = this.deltaSyncService.calculateDelta(
            currentVersion,
            drawing
          );

          if (!delta) {
            XLogger.warn(`No delta calculated for drawing: ${drawing.name}`);
            // If no delta, just save the entire drawing
            await MicrosoftFetchService.saveFile(
              MicrosoftProvider.SYNC_FOLDER_NAME,
              `${drawing.name}.json`,
              drawing,
              accessToken
            );
          } else {
            const updatedDrawing = this.deltaSyncService.applyDelta(
              currentVersion,
              delta
            );

            if (!updatedDrawing) {
              XLogger.warn(
                `Failed to apply delta for drawing: ${drawing.name}`
              );
              // If delta application fails, just save the entire drawing
              await MicrosoftFetchService.saveFile(
                MicrosoftProvider.SYNC_FOLDER_NAME,
                `${drawing.name}.json`,
                drawing,
                accessToken
              );
            } else {
              // Update sync status progress
              this.deltaSyncService.updateSyncStatus(drawing.id, {
                status: SyncStatus.SYNCING,
                version: drawing.data.versionFiles,
                progress: 50,
              });

              await MicrosoftFetchService.saveFile(
                MicrosoftProvider.SYNC_FOLDER_NAME,
                `${drawing.name}.json`,
                updatedDrawing,
                accessToken
              );
            }
          }
        } catch (deltaError) {
          XLogger.error(
            `Error calculating or applying delta for drawing ${drawing.name}: ${
              deltaError instanceof Error
                ? deltaError.message
                : String(deltaError)
            }`
          );
          // If delta calculation or application fails, just save the entire drawing
          await MicrosoftFetchService.saveFile(
            MicrosoftProvider.SYNC_FOLDER_NAME,
            `${drawing.name}.json`,
            drawing,
            accessToken
          );
        }
      } else {
        // New file, save entire drawing
        await MicrosoftFetchService.saveFile(
          MicrosoftProvider.SYNC_FOLDER_NAME,
          `${drawing.name}.json`,
          drawing,
          accessToken
        );
      }

      // Update sync status to synced
      this.deltaSyncService.updateSyncStatus(drawing.id, {
        status: SyncStatus.SYNCED,
        version: drawing.data.versionFiles,
        lastSyncedAt: new Date(),
        progress: 100,
      });
    } catch (error) {
      const operation = isUpdate ? "updating" : "saving";
      XLogger.error(
        `Error ${operation} drawing to Microsoft OneDrive: ${
          error instanceof Error ? error.message : String(error)
        }`
      );

      // Update sync status to error
      this.deltaSyncService.updateSyncStatus(drawing.id, {
        status: SyncStatus.ERROR,
        version: drawing.data.versionFiles,
        error: error instanceof Error ? error.message : String(error),
      });

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

      // Update sync status to syncing
      this.deltaSyncService.updateSyncStatus(drawingName, {
        status: SyncStatus.SYNCING,
        version: "",
        progress: 0,
      });

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

      // Clear sync status after successful deletion
      this.deltaSyncService.clearSyncStatus(drawingName);
    } catch (error) {
      XLogger.error(
        `Error deleting drawing from Microsoft OneDrive: ${
          error instanceof Error ? error.message : String(error)
        }`
      );

      // Update sync status to error
      this.deltaSyncService.updateSyncStatus(drawingName, {
        status: SyncStatus.ERROR,
        version: "",
        error: error instanceof Error ? error.message : String(error),
      });

      throw error;
    }
  }

  /**
   * Sync files between local storage and Microsoft OneDrive
   */
  public async syncFiles(): Promise<void> {
    try {
      XLogger.info("Starting sync from OneDrive to local storage");

      // Update sync status for all files to pending
      const folders = await browser.storage.local.get("folders");
      const oneDriveFolder = folders.folders?.find(
        (f: any) => f.name === MicrosoftProvider.SYNC_FOLDER_NAME
      );

      if (oneDriveFolder?.drawingIds) {
        for (const drawingId of oneDriveFolder.drawingIds) {
          this.deltaSyncService.updateSyncStatus(drawingId, {
            status: SyncStatus.PENDING,
            version: "",
            progress: 0,
          });
        }
      }

      const oneDriveDrawings = await this.getOneDriveFiles();
      XLogger.info(
        `Retrieved ${oneDriveDrawings.length} drawings from OneDrive`
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

        // Update sync status to syncing
        this.deltaSyncService.updateSyncStatus(drawing.id, {
          status: SyncStatus.SYNCING,
          version: drawing.data.versionFiles,
          progress: 0,
        });

        const localDrawing = (await browser.storage.local.get(drawing.id))[
          drawing.id
        ];

        // Check for conflicts by comparing only the relevant drawing data
        if (localDrawing) {
          // Skip conflict check if the drawing is already in conflict state
          const currentStatus = this.deltaSyncService.getSyncStatus(drawing.id);
          if (currentStatus?.status === SyncStatus.CONFLICT) {
            XLogger.info(
              `Skipping conflict check for drawing in conflict state: ${drawing.name}`
            );
            continue;
          }

          const hasConflict = this.deltaSyncService.hasDrawingConflict(
            localDrawing,
            drawing
          );
          if (hasConflict) {
            XLogger.info(`Conflict detected for drawing: ${drawing.name}`);

            // Update sync status to conflict
            this.deltaSyncService.updateSyncStatus(drawing.id, {
              status: SyncStatus.CONFLICT,
              version: drawing.data.versionFiles,
            });

            // Emit conflict event for UI to handle
            try {
              await browser.runtime.sendMessage({
                type: "MICROSOFT_SYNC_CONFLICT",
                payload: {
                  drawingId: drawing.id,
                  localDrawing,
                  oneDriveDrawing: drawing,
                },
              });
            } catch (error) {
              // If the message can't be sent (e.g., popup not open), just log it
              XLogger.warn(
                `Could not send sync conflict message for drawing ${
                  drawing.name
                }: ${error instanceof Error ? error.message : String(error)}`
              );
            }
            continue;
          }
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

        // Update sync status to synced
        this.deltaSyncService.updateSyncStatus(drawing.id, {
          status: SyncStatus.SYNCED,
          version: drawing.data.versionFiles,
          lastSyncedAt: new Date(),
          progress: 100,
        });
      }

      // Save updated folder
      XLogger.info("Saving updated OneDrive folder to local storage");
      await browser.storage.local.set({ folders: folders.folders });
      XLogger.info("Sync from OneDrive to local storage completed");
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      XLogger.error(`Error syncing OneDrive files: ${errorMessage}`);

      // Update sync status for all files to error
      const folders = await browser.storage.local.get("folders");
      const oneDriveFolder = folders.folders?.find(
        (f: any) => f.name === MicrosoftProvider.SYNC_FOLDER_NAME
      );

      if (oneDriveFolder?.drawingIds) {
        for (const drawingId of oneDriveFolder.drawingIds) {
          this.deltaSyncService.updateSyncStatus(drawingId, {
            status: SyncStatus.ERROR,
            version: "",
            error: errorMessage,
          });
        }
      }

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
          // Continue with other files even if one fails
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

  /**
   * Check if there is a conflict between two drawings by comparing only the relevant data
   */
  private hasDrawingConflict(
    localDrawing: IDrawing,
    oneDriveDrawing: IDrawing
  ): boolean {
    try {
      // Parse and normalize the drawing data
      const parseDrawingData = (drawing: IDrawing) => {
        try {
          // First parse attempt
          let data =
            typeof drawing.data === "string"
              ? JSON.parse(drawing.data)
              : drawing.data;

          // Second parse if needed (for double-stringified data)
          if (typeof data === "string") {
            data = JSON.parse(data);
          }

          // Ensure we have an array of elements
          return Array.isArray(data) ? data : [];
        } catch (error) {
          XLogger.error("Error parsing drawing data:", error);
          return [];
        }
      };

      const localElements = parseDrawingData(localDrawing);
      const oneDriveElements = parseDrawingData(oneDriveDrawing);

      // Log the full data for comparison
      XLogger.info("Drawing comparison:", {
        local: {
          name: localDrawing.name,
          id: localDrawing.id,
          elementCount: localElements.length,
        },
        oneDrive: {
          name: oneDriveDrawing.name,
          id: oneDriveDrawing.id,
          elementCount: oneDriveElements.length,
        },
      });

      // Compare elements if they exist
      if (Array.isArray(localElements) && Array.isArray(oneDriveElements)) {
        // Log element counts
        XLogger.info("Element counts:", {
          local: localElements.length,
          oneDrive: oneDriveElements.length,
        });

        // Create maps for easier comparison
        const localMap = new Map<string, any>(
          localElements.map((e: any) => [e.id, e])
        );
        const oneDriveMap = new Map<string, any>(
          oneDriveElements.map((e: any) => [e.id, e])
        );

        // Find differences
        const differences = [];

        // Check for elements in local but not in OneDrive
        for (const [id, localElement] of Array.from(localMap.entries())) {
          const oneDriveElement = oneDriveMap.get(id);
          if (!oneDriveElement) {
            differences.push({
              type: "missing_in_onedrive",
              id,
              element: localElement,
            });
            continue;
          }

          // Compare properties
          const props = this.getEssentialPropsForType(localElement.type || "");
          const propertyDiffs = [];

          for (const prop of props) {
            const localValue = (localElement as any)[prop];
            const oneDriveValue = (oneDriveElement as any)[prop];

            if (JSON.stringify(localValue) !== JSON.stringify(oneDriveValue)) {
              propertyDiffs.push({
                property: prop,
                local: localValue,
                oneDrive: oneDriveValue,
              });
            }
          }

          if (propertyDiffs.length > 0) {
            differences.push({
              type: "different_properties",
              id,
              differences: propertyDiffs,
            });
          }
        }

        // Check for elements in OneDrive but not in local
        for (const [id, oneDriveElement] of Array.from(oneDriveMap.entries())) {
          if (!localMap.has(id)) {
            differences.push({
              type: "missing_in_local",
              id,
              element: oneDriveElement,
            });
          }
        }

        // Log all differences in a readable format
        if (differences.length > 0) {
          XLogger.info(
            "Found differences:",
            JSON.stringify(differences, null, 2)
          );
          return true;
        }
      } else {
        // Log if elements are missing or not arrays
        XLogger.info("Invalid elements data:", {
          localIsArray: Array.isArray(localElements),
          oneDriveIsArray: Array.isArray(oneDriveElements),
        });
        return true;
      }

      return false;
    } catch (error) {
      XLogger.error("Error checking for drawing conflicts:", error);
      return true;
    }
  }

  /**
   * Get the essential properties to compare for each element type
   */
  private getEssentialPropsForType(type: string): string[] {
    // Common properties that affect visual appearance for all element types
    const commonProps = [
      "type",
      "x",
      "y",
      "width",
      "height",
      "angle",
      "strokeColor",
      "backgroundColor",
      "fillStyle",
      "strokeWidth",
      "strokeStyle",
      "roughness",
      "opacity",
      "isDeleted",
    ];

    // Type-specific properties that affect visual appearance
    switch (type) {
      case "text":
        return [
          ...commonProps,
          "text",
          "fontSize",
          "fontFamily",
          "textAlign",
          "verticalAlign",
          "baseline",
        ];
      case "image":
        return [...commonProps, "scaleX", "scaleY", "status", "fileId"];
      case "line":
      case "arrow":
        return [
          ...commonProps,
          "startX",
          "startY",
          "endX",
          "endY",
          "startArrowhead",
          "endArrowhead",
          "label",
          "labelPosition",
        ];
      case "freedraw":
        return [...commonProps, "points"];
      case "frame":
        return [...commonProps, "name", "background"];
      default:
        return commonProps;
    }
  }

  /**
   * Sync a single file to Microsoft OneDrive
   */
  public async syncFile(drawingId: string): Promise<void> {
    try {
      XLogger.info(`Starting sync of drawing ${drawingId} to OneDrive`);

      // Get drawing from local storage
      const drawing = (await browser.storage.local.get(drawingId))[drawingId];
      if (!drawing) {
        throw new Error(`Drawing ${drawingId} not found in local storage`);
      }

      // Update sync status to syncing
      this.deltaSyncService.updateSyncStatus(drawingId, {
        status: SyncStatus.SYNCING,
        version: drawing.data.versionFiles,
        progress: 0,
      });

      // Get OneDrive file
      const oneDriveFile = await this.getOneDriveFile(drawingId);
      if (!oneDriveFile) {
        throw new Error(`Drawing ${drawingId} not found in OneDrive`);
      }

      // Check for conflicts
      const hasConflict = this.deltaSyncService.hasDrawingConflict(
        drawing,
        oneDriveFile
      );
      if (hasConflict) {
        XLogger.info(`Conflict detected for drawing: ${drawing.name}`);

        // Update sync status to conflict
        this.deltaSyncService.updateSyncStatus(drawingId, {
          status: SyncStatus.CONFLICT,
          version: oneDriveFile.data.versionFiles,
        });

        // Emit conflict event for UI to handle
        try {
          await browser.runtime.sendMessage({
            type: "MICROSOFT_SYNC_CONFLICT",
            payload: {
              drawingId,
              localDrawing: drawing,
              oneDriveDrawing: oneDriveFile,
            },
          });
        } catch (error) {
          // If the message can't be sent (e.g., popup not open), just log it
          XLogger.warn(
            `Could not send sync conflict message for drawing ${
              drawing.name
            }: ${error instanceof Error ? error.message : String(error)}`
          );
        }
        return;
      }

      // Upload to OneDrive
      await this.uploadToOneDrive(drawing);

      // Update sync status to synced
      this.deltaSyncService.updateSyncStatus(drawingId, {
        status: SyncStatus.SYNCED,
        version: drawing.data.versionFiles,
        lastSyncedAt: new Date(),
        progress: 100,
      });

      XLogger.info(`Sync of drawing ${drawingId} to OneDrive completed`);
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      XLogger.error(
        `Error syncing drawing ${drawingId} to OneDrive: ${errorMessage}`
      );

      // Update sync status to error
      this.deltaSyncService.updateSyncStatus(drawingId, {
        status: SyncStatus.ERROR,
        version: "",
        error: errorMessage,
      });

      throw error;
    }
  }

  /**
   * Get a single file from OneDrive
   */
  private async getOneDriveFile(drawingId: string): Promise<IDrawing | null> {
    try {
      const files = await this.getOneDriveFiles();
      return files.find((file) => file.id === drawingId) || null;
    } catch (error) {
      XLogger.error(
        `Error getting OneDrive file ${drawingId}: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      return null;
    }
  }

  /**
   * Upload a drawing to OneDrive
   */
  private async uploadToOneDrive(drawing: IDrawing): Promise<void> {
    try {
      const filePath = `${MicrosoftProvider.SYNC_FOLDER_NAME}/${drawing.id}.excalidraw`;
      const content = JSON.stringify(drawing.data);

      await MicrosoftFetchService.saveFile(
        MicrosoftProvider.SYNC_FOLDER_NAME,
        `${drawing.name}.json`,
        drawing,
        await this.authService.getAccessToken()
      );

      XLogger.info(`Successfully uploaded drawing ${drawing.id} to OneDrive`);
    } catch (error) {
      XLogger.error(
        `Error uploading drawing ${drawing.id} to OneDrive: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }
}
