import { IDrawing } from "../../interfaces/drawing.interface";

/**
 * Interface for sync providers (e.g., Microsoft OneDrive, Google Drive)
 */
export interface SyncProvider {
  /**
   * Initialize the sync provider
   */
  initialize(): Promise<void>;

  /**
   * Check if the user is authenticated with the provider
   */
  isAuthenticated(): Promise<boolean>;

  /**
   * Save a drawing to the cloud
   */
  saveDrawing(drawing: IDrawing): Promise<void>;

  /**
   * Update an existing drawing in the cloud
   */
  updateDrawing(drawing: IDrawing): Promise<void>;

  /**
   * Delete a drawing from the cloud
   */
  deleteDrawing(drawingName: string): Promise<void>;

  /**
   * Sync files between local and cloud
   */
  syncFiles(): Promise<void>;
}
