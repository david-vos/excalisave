import { IDrawing } from "../../interfaces/drawing.interface";
import {
  FileDelta,
  SyncStatus,
  SyncStatusInfo,
} from "../../interfaces/sync-status.interface";
import { XLogger } from "../../lib/logger";

interface DrawingElement {
  id: string;
  [key: string]: any;
}

interface DrawingData {
  elements: DrawingElement[];
  versionFiles: string;
  excalidraw: string;
  excalidrawState: string;
  versionDataState: string;
  [key: string]: any;
}

interface ExtendedDrawing extends IDrawing {
  data: DrawingData;
}

export class DeltaSyncService {
  private static instance: DeltaSyncService;
  private syncStatusMap: Map<string, SyncStatusInfo> = new Map();

  private constructor() {}

  public static getInstance(): DeltaSyncService {
    if (!DeltaSyncService.instance) {
      DeltaSyncService.instance = new DeltaSyncService();
    }
    return DeltaSyncService.instance;
  }

  /**
   * Parse and normalize drawing data
   */
  private parseDrawingData(drawing: IDrawing): DrawingElement[] {
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
  }

  /**
   * Get essential properties to compare for each element type
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
   * Check if there is a conflict between two drawings
   */
  public hasDrawingConflict(
    localDrawing: IDrawing,
    oneDriveDrawing: IDrawing
  ): boolean {
    try {
      const localElements = this.parseDrawingData(localDrawing);
      const oneDriveElements = this.parseDrawingData(oneDriveDrawing);

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
   * Calculate delta between two versions of a drawing
   */
  public calculateDelta(
    oldDrawing: ExtendedDrawing,
    newDrawing: ExtendedDrawing
  ): FileDelta | null {
    // Validate that both drawings have the required data structure
    if (!oldDrawing?.data?.elements || !newDrawing?.data?.elements) {
      XLogger.warn("Cannot calculate delta: missing elements in drawing data");
      return null;
    }

    const delta: FileDelta = {
      path: newDrawing.id,
      changes: {
        added: [],
        modified: [],
        deleted: [],
      },
      baseVersion: oldDrawing.data.versionFiles || "",
      targetVersion: newDrawing.data.versionFiles || "",
    };

    // Compare elements
    const oldElements = new Set(
      oldDrawing.data.elements.map((el: DrawingElement) => el.id)
    );
    const newElements = new Set(
      newDrawing.data.elements.map((el: DrawingElement) => el.id)
    );

    // Find added and deleted elements
    for (const elementId of Array.from(newElements)) {
      if (!oldElements.has(elementId)) {
        delta.changes.added.push(elementId);
      }
    }

    for (const elementId of Array.from(oldElements)) {
      if (!newElements.has(elementId)) {
        delta.changes.deleted.push(elementId);
      }
    }

    // Find modified elements
    const oldElementsMap = new Map(
      oldDrawing.data.elements.map((el: DrawingElement) => [el.id, el])
    );
    const newElementsMap = new Map(
      newDrawing.data.elements.map((el: DrawingElement) => [el.id, el])
    );

    for (const [elementId, newElement] of Array.from(
      newElementsMap.entries()
    )) {
      const oldElement = oldElementsMap.get(elementId);
      if (
        oldElement &&
        JSON.stringify(oldElement) !== JSON.stringify(newElement)
      ) {
        delta.changes.modified.push(elementId);
      }
    }

    return delta;
  }

  /**
   * Apply delta to a drawing
   */
  public applyDelta(
    drawing: ExtendedDrawing,
    delta: FileDelta
  ): ExtendedDrawing | null {
    // Validate that the drawing has the required data structure
    if (!drawing?.data?.elements) {
      XLogger.warn("Cannot apply delta: missing elements in drawing data");
      return null;
    }

    const updatedDrawing = { ...drawing };
    const elementsMap = new Map(
      drawing.data.elements.map((el: DrawingElement) => [el.id, el])
    );

    // Apply changes
    for (const elementId of delta.changes.added) {
      const newElement = drawing.data.elements.find(
        (el: DrawingElement) => el.id === elementId
      );
      if (newElement) {
        elementsMap.set(elementId, newElement);
      }
    }

    for (const elementId of delta.changes.deleted) {
      elementsMap.delete(elementId);
    }

    for (const elementId of delta.changes.modified) {
      const newElement = drawing.data.elements.find(
        (el: DrawingElement) => el.id === elementId
      );
      if (newElement) {
        elementsMap.set(elementId, newElement);
      }
    }

    updatedDrawing.data.elements = Array.from(elementsMap.values());
    updatedDrawing.data.versionFiles = delta.targetVersion;

    return updatedDrawing;
  }

  /**
   * Update sync status for a drawing
   */
  public updateSyncStatus(drawingId: string, status: SyncStatusInfo): void {
    this.syncStatusMap.set(drawingId, status);
    XLogger.info(
      `Updated sync status for drawing ${drawingId}: ${status.status}`
    );
  }

  /**
   * Get sync status for a drawing
   */
  public getSyncStatus(drawingId: string): SyncStatusInfo | undefined {
    return this.syncStatusMap.get(drawingId);
  }

  /**
   * Clear sync status for a drawing
   */
  public clearSyncStatus(drawingId: string): void {
    this.syncStatusMap.delete(drawingId);
  }
}
