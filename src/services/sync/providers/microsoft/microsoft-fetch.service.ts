import { XLogger } from "../../../../lib/logger";

/**
 * Service for making API calls to Microsoft Graph API
 */
export class MicrosoftFetchService {
  private static readonly API_BASE_URL = "https://graph.microsoft.com/v1.0";

  /**
   * Makes an authenticated request to the Microsoft Graph API
   */
  public static async fetch(
    endpoint: string,
    options: RequestInit = {},
    accessToken: string
  ): Promise<Response> {
    // Construct the full URL - all endpoints are relative to the base URL
    const url = `${this.API_BASE_URL}${
      endpoint.startsWith("/") ? endpoint : `/${endpoint}`
    }`;

    XLogger.info(`Making request to Microsoft API: ${url}`);
    XLogger.info(`Request method: ${options.method || "GET"}`);

    const headers = {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...options.headers,
    };

    try {
      const response = await fetch(url, {
        ...options,
        headers,
      });

      XLogger.info(
        `Response status: ${response.status} ${response.statusText}`
      );

      if (!response.ok) {
        const errorText = await response.text();
        XLogger.error(
          `Microsoft API request failed: ${response.statusText} - ${errorText}`
        );
        throw new Error(
          `Microsoft API request failed: ${response.statusText} - ${errorText}`
        );
      }

      return response;
    } catch (error) {
      XLogger.error(`Error fetching from Microsoft API: ${endpoint}`, error);
      throw error;
    }
  }

  /**
   * Gets the content of a file from OneDrive by ID
   */
  public static async getFileContent(
    fileId: string,
    accessToken: string
  ): Promise<any> {
    XLogger.info(`Getting content for file with ID: ${fileId}`);
    const response = await this.fetch(
      `/me/drive/items/${fileId}/content`,
      {},
      accessToken
    );

    XLogger.info(`Retrieved content for file with ID: ${fileId}`);
    // Since we know we're dealing with JSON files, we can parse the text content
    const text = await response.text();
    XLogger.info(`Parsed text content for file with ID: ${fileId}`);

    try {
      // Use a more robust JSON parsing process
      const json = JSON.parse(text, (key, value) => {
        // Handle special cases for data types
        if (value === null && key === "") {
          // This is the root object, return as is
          return value;
        }
        return value;
      });

      XLogger.info(`Successfully parsed JSON for file with ID: ${fileId}`);
      return json;
    } catch (error) {
      XLogger.error(`Error parsing JSON for file with ID: ${fileId}`, error);
      throw error;
    }
  }

  /**
   * Gets the content of a file from OneDrive by path
   */
  public static async getFileContentByPath(
    path: string,
    accessToken: string
  ): Promise<any> {
    XLogger.info(`Getting content for file with path: ${path}`);
    const response = await this.fetch(
      `/me/drive/root:/${path}:/content`,
      {},
      accessToken
    );

    XLogger.info(`Retrieved content for file with path: ${path}`);
    // Since we know we're dealing with JSON files, we can parse the text content
    const text = await response.text();
    XLogger.info(`Parsed text content for file with path: ${path}`);

    try {
      const json = JSON.parse(text);
      XLogger.info(`Successfully parsed JSON for file with path: ${path}`);
      return json;
    } catch (error) {
      XLogger.error(`Error parsing JSON for file with path: ${path}`, error);
      throw error;
    }
  }

  /**
   * Checks if a folder exists in OneDrive
   */
  public static async checkFolderExists(
    folderName: string,
    accessToken: string
  ): Promise<Response> {
    return this.fetch(`/me/drive/root:/${folderName}`, {}, accessToken);
  }

  /**
   * Creates a folder in OneDrive
   */
  public static async createFolder(
    folderName: string,
    accessToken: string
  ): Promise<Response> {
    return this.fetch(
      "/me/drive/root/children",
      {
        method: "POST",
        body: JSON.stringify({
          name: folderName,
          folder: {},
        }),
      },
      accessToken
    );
  }

  /**
   * Gets all items in a folder from OneDrive
   */
  public static async getFolderItems(
    folderId: string,
    accessToken: string
  ): Promise<Response> {
    return this.fetch(`/me/drive/items/${folderId}/children`, {}, accessToken);
  }

  /**
   * Saves or updates a file in OneDrive
   */
  public static async saveFile(
    folderName: string,
    fileName: string,
    content: any,
    accessToken: string
  ): Promise<Response> {
    // Use a more robust JSON stringification process
    const jsonString = JSON.stringify(
      content,
      (_, value) => {
        // Handle special cases for data types
        if (value === undefined) {
          return null; // Convert undefined to null
        }
        if (typeof value === "function") {
          return null; // Skip functions
        }
        return value;
      },
      0
    ); // Use 0 spaces for compact JSON

    return this.fetch(
      `/me/drive/root:/${folderName}/${fileName}:/content`,
      {
        method: "PUT",
        body: jsonString,
      },
      accessToken
    );
  }

  /**
   * Deletes a file from OneDrive
   */
  public static async deleteFile(
    folderName: string,
    fileName: string,
    accessToken: string
  ): Promise<Response> {
    XLogger.info(`Deleting file from OneDrive: ${folderName}/${fileName}`);
    const response = await this.fetch(
      `/me/drive/root:/${folderName}/${fileName}`,
      {
        method: "DELETE",
      },
      accessToken
    );
    XLogger.info(
      `Delete file response status: ${response.status} ${response.statusText}`
    );
    return response;
  }
}
