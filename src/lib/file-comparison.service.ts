import { IDrawing } from "../interfaces/drawing.interface";
import { MicrosoftApiService } from "./microsoft-api.service";
import { XLogger } from "./logger";

export class FileComparisonService {
  private static instance: FileComparisonService;

  private constructor() {}

  public static getInstance(): FileComparisonService {
    if (!FileComparisonService.instance) {
      FileComparisonService.instance = new FileComparisonService();
    }
    return FileComparisonService.instance;
  }

  public static async mergeFiles(drawingId: string): Promise<void> {
    if (!drawingId) {
      throw new Error("drawingId is required");
    }

    try {
      const { localFile, oneDriveFile } =
        await MicrosoftApiService.getCompareFiles(drawingId);

      // Validate files
      if (!localFile && !oneDriveFile) {
        XLogger.error("Both local and OneDrive files are null", { drawingId });
        throw new Error("Both files are null - cannot merge");
      }

      // If only one file exists, validate and use that one
      if (!localFile || !oneDriveFile) {
        const fileToUse = localFile || oneDriveFile;
        if (!fileToUse) {
          XLogger.error("No valid file to use", { drawingId });
          throw new Error("No valid file to use");
        }

        // Validate file structure
        if (!fileToUse.data || !fileToUse.data.excalidraw) {
          XLogger.error("Invalid file structure", {
            drawingId,
            file: fileToUse,
          });
          throw new Error("Invalid file structure");
        }

        await MicrosoftApiService.saveDrawing(fileToUse);
        return;
      }

      // Validate both files have required structure
      if (!localFile.data?.excalidraw || !oneDriveFile.data?.excalidraw) {
        XLogger.error("Invalid file structure", {
          drawingId,
          localFile: !!localFile.data?.excalidraw,
          oneDriveFile: !!oneDriveFile.data?.excalidraw,
        });
        throw new Error("Invalid file structure");
      }

      // Parse the excalidraw data from both files
      let localElements, oneDriveElements;
      try {
        localElements = JSON.parse(localFile.data.excalidraw);
        oneDriveElements = JSON.parse(oneDriveFile.data.excalidraw);
      } catch (error) {
        XLogger.error("Error parsing excalidraw data", { drawingId, error });
        throw new Error("Invalid excalidraw data format");
      }

      // Validate parsed elements are arrays
      if (!Array.isArray(localElements) || !Array.isArray(oneDriveElements)) {
        XLogger.error("Invalid elements format", {
          drawingId,
          localIsArray: Array.isArray(localElements),
          oneDriveIsArray: Array.isArray(oneDriveElements),
        });
        throw new Error("Invalid elements format");
      }

      // Create a map of elements by ID for both files
      const localElementsMap = new Map<string, any>(
        Array.from(localElements)
          .map((el: any) => {
            if (!el?.id) {
              XLogger.warn("Element missing ID", { drawingId, element: el });
              return null;
            }
            return [el.id, el] as [string, any];
          })
          .filter((entry): entry is [string, any] => entry !== null)
      );

      const oneDriveElementsMap = new Map<string, any>(
        Array.from(oneDriveElements)
          .map((el: any) => {
            if (!el?.id) {
              XLogger.warn("Element missing ID", { drawingId, element: el });
              return null;
            }
            return [el.id, el] as [string, any];
          })
          .filter((entry): entry is [string, any] => entry !== null)
      );

      // Get timestamps for conflict resolution
      const localTimestamp = new Date(localFile.createdAt).getTime();
      const oneDriveTimestamp = new Date(oneDriveFile.createdAt).getTime();

      if (isNaN(localTimestamp) || isNaN(oneDriveTimestamp)) {
        XLogger.error("Invalid timestamp", {
          drawingId,
          localTimestamp,
          oneDriveTimestamp,
        });
        throw new Error("Invalid timestamp");
      }

      // Merge elements - keep all unique elements and resolve conflicts by timestamp
      const mergedElements = [];
      const allIds = new Set([
        ...Array.from(localElementsMap.keys()),
        ...Array.from(oneDriveElementsMap.keys()),
      ]);

      for (const id of Array.from(allIds)) {
        const localElement = localElementsMap.get(id);
        const oneDriveElement = oneDriveElementsMap.get(id);

        if (!localElement) {
          // Element only exists in OneDrive
          mergedElements.push(oneDriveElement);
        } else if (!oneDriveElement) {
          // Element only exists locally
          mergedElements.push(localElement);
        } else {
          // Element exists in both - use the one with the latest timestamp
          const elementLocalTimestamp = new Date(
            (localElement as any).updated || localFile.createdAt
          ).getTime();
          const elementOneDriveTimestamp = new Date(
            (oneDriveElement as any).updated || oneDriveFile.createdAt
          ).getTime();

          if (isNaN(elementLocalTimestamp) || isNaN(elementOneDriveTimestamp)) {
            XLogger.warn("Invalid element timestamp, using file timestamp", {
              drawingId,
              elementId: id,
            });
            // Fallback to file timestamps if element timestamps are invalid
            mergedElements.push(
              oneDriveTimestamp > localTimestamp
                ? oneDriveElement
                : localElement
            );
          } else {
            // Use the element with the most recent timestamp
            mergedElements.push(
              elementOneDriveTimestamp > elementLocalTimestamp
                ? oneDriveElement
                : localElement
            );
          }
        }
      }

      // Validate we have elements to save
      if (mergedElements.length === 0) {
        XLogger.error("No elements to save after merge", { drawingId });
        throw new Error("No elements to save after merge");
      }

      // Create the merged file with the latest timestamp
      const mergedFile: IDrawing = {
        ...localFile,
        data: {
          ...localFile.data,
          excalidraw: JSON.stringify(mergedElements),
          // Increment version to indicate a merge occurred
          versionFiles: (
            Math.max(
              parseInt(localFile.data.versionFiles) || 0,
              parseInt(oneDriveFile.data.versionFiles) || 0
            ) + 1
          ).toString(),
        },
        createdAt: new Date().toISOString(),
      };

      await MicrosoftApiService.saveDrawing(mergedFile);
    } catch (error) {
      XLogger.error("Error merging files", { drawingId, error });
      throw error;
    }
  }
}
