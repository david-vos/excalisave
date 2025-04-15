import { XLogger } from "./logger";

export class MicrosoftFetchService {
  private static readonly API_BASE_URL = "https://graph.microsoft.com/v1.0";
  private static readonly SYNC_FOLDER_NAME = "ExcaliSave";

  /**
   * Makes an authenticated request to the Microsoft Graph API
   */
  public static async fetch(
    endpoint: string,
    options: RequestInit = {},
    accessToken: string
  ): Promise<Response> {
    const url = endpoint.startsWith("http")
      ? endpoint
      : `${this.API_BASE_URL}${
          endpoint.startsWith("/") ? endpoint : `/${endpoint}`
        }`;

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

      if (!response.ok) {
        throw new Error(`Microsoft API request failed: ${response.statusText}`);
      }

      return response;
    } catch (error) {
      XLogger.error(`Error fetching from Microsoft API: ${endpoint}`, error);
      throw error;
    }
  }

  /**
   * Gets the content of a file from OneDrive
   */
  public static async getFileContent(
    fileId: string,
    accessToken: string
  ): Promise<any> {
    const response = await this.fetch(
      `/me/drive/items/${fileId}/content`,
      {},
      accessToken
    );
    return response.json();
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
    return this.fetch(
      `/me/drive/root:/${folderName}/${fileName}:/content`,
      {
        method: "PUT",
        body: JSON.stringify(content),
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
    return this.fetch(
      `/me/drive/root:/${folderName}/${fileName}`,
      {
        method: "DELETE",
      },
      accessToken
    );
  }

  /**
   * Exchanges an authorization code for an access token
   */
  public static async exchangeCodeForToken(
    code: string,
    redirectUri: string,
    codeVerifier: string,
    clientId: string,
    tenantId: string = "consumers"
  ): Promise<string> {
    try {
      const response = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: new URLSearchParams({
            client_id: clientId,
            code: code,
            redirect_uri: redirectUri,
            grant_type: "authorization_code",
            code_verifier: codeVerifier,
          }),
        }
      );

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(
          `Failed to exchange code for token: ${
            errorData.error_description || response.statusText
          }`
        );
      }

      const data = await response.json();
      return data.access_token;
    } catch (error) {
      XLogger.error("Error exchanging code for token", error);
      throw error;
    }
  }
}
