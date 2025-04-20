/**
 * Drawing interface
 *
 * This is the how the drawing is stored in browser storage
 */
export interface DrawingElement {
  id: string;
  [key: string]: any;
}

export interface DrawingData {
  elements: DrawingElement[];
  excalidraw: string;
  excalidrawState: string;
  versionFiles: string;
  versionDataState: string;
  [key: string]: any;
}

export interface IDrawing {
  id: string;
  name: string;
  createdAt: string;
  updatedAt: string;
  imageBase64?: string;
  viewBackgroundColor?: string;
  data: DrawingData;
}
