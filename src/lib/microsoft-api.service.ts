import { IDrawing } from "../interfaces/drawing.interface";
import { XLogger } from "./logger";
import { browser } from "webextension-polyfill-ts";

export class MicrosoftApiService {
  private static instance: MicrosoftApiService;
  public static accessToken: string | null = null;
  public static initialized = false;
  public static readonly TENANT_ID = "common";
  public static readonly CLIENT_ID = "YOUR_CLIENT_ID";
  public static readonly REDIRECT_URI = "https://excalidraw.com/auth-callback";
  public static readonly SCOPES = ["Files.ReadWrite", "User.Read"];
  private static readonly GRAPH_API_ENDPOINT =
    "https://graph.microsoft.com/v1.0";

  private static readonly API_BASE_URL = "https://graph.microsoft.com/v1.0";

  private constructor() {}

  public static getInstance(): MicrosoftApiService {
    if (!MicrosoftApiService.instance) {
      MicrosoftApiService.instance = new MicrosoftApiService();
    }
    return MicrosoftApiService.instance;
  }

  public static async initialize(): Promise<void> {
    if (MicrosoftApiService.initialized) {
      return;
    }
    try {
      const result = await browser.storage.local.get("microsoftAccessToken");
      if (result.microsoftAccessToken) {
        MicrosoftApiService.accessToken = result.microsoftAccessToken;
      } else {
        MicrosoftApiService.accessToken = null;
        MicrosoftApiService.initialized = false;
      }
    } catch (error) {
      XLogger.error("Error initializing Microsoft API", error);
    }
    MicrosoftApiService.initialized = true;
  }

  public async saveAccessToken(token: string): Promise<void> {
    try {
      await browser.storage.local.set({ microsoftAccessToken: token });
      MicrosoftApiService.accessToken = token;
      MicrosoftApiService.initialized = true;
    } catch (error) {
      XLogger.error("Error saving Microsoft access token", error);
      throw error;
    }
  }

  public async removeAccessToken(): Promise<void> {
    try {
      await browser.storage.local.remove("microsoftAccessToken");
      MicrosoftApiService.accessToken = null;
      MicrosoftApiService.initialized = false;
    } catch (error) {
      XLogger.error("Error removing Microsoft access token", error);
      throw error;
    }
  }

  public isAuthenticated(): boolean {
    return (
      MicrosoftApiService.initialized &&
      MicrosoftApiService.accessToken !== null
    );
  }

  private async getAccessToken(): Promise<string> {
    if (!MicrosoftApiService.accessToken) {
      throw new Error("Not authenticated with Microsoft");
    }
    return MicrosoftApiService.accessToken;
  }

  public async saveToOneDrive(
    fileName: string,
    content: string
  ): Promise<string> {
    try {
      const token = await this.getAccessToken();
      const response = await fetch(
        `${MicrosoftApiService.GRAPH_API_ENDPOINT}/me/drive/root:/${fileName}:/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: content,
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to save to OneDrive: ${response.statusText}`);
      }

      const data = await response.json();
      return data.webUrl;
    } catch (error) {
      XLogger.error("Error saving to OneDrive", error);
      throw error;
    }
  }

  public async loadFromOneDrive(fileId: string): Promise<string> {
    try {
      const token = await this.getAccessToken();
      const response = await fetch(
        `${MicrosoftApiService.GRAPH_API_ENDPOINT}/me/drive/items/${fileId}/content`,
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );
      if (!response.ok) {
        throw new Error(`Failed to load from OneDrive: ${response.statusText}`);
      }

      return await response.text();
    } catch (error) {
      XLogger.error("Error loading from OneDrive", error);
      throw error;
    }
  }

  static async saveDrawing(drawing: IDrawing): Promise<void> {
    if (!MicrosoftApiService.accessToken) {
      await MicrosoftApiService.initialize();
    }

    try {
      const response = await fetch(
        `${MicrosoftApiService.API_BASE_URL}/me/drive/root:/ExcaliSave/${drawing.name}.json:/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${MicrosoftApiService.accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(drawing),
        }
      );

      if (!response.ok) {
        if (response.status === 401) {
          // Token expired, try to reinitialize
          await MicrosoftApiService.initialize();
          return this.saveDrawing(drawing);
        }
        throw new Error(`Failed to save drawing: ${response.statusText}`);
      }

      XLogger.log("Drawing saved to Microsoft OneDrive successfully");
    } catch (error) {
      XLogger.error("Error saving drawing to Microsoft OneDrive", error);
      throw error;
    }
  }

  static async updateDrawing(drawing: IDrawing): Promise<void> {
    if (!MicrosoftApiService.accessToken) {
      await MicrosoftApiService.initialize();
    }

    try {
      const response = await fetch(
        `${MicrosoftApiService.API_BASE_URL}/me/drive/root:/ExcaliSave/${drawing.name}.json:/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${MicrosoftApiService.accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(drawing),
        }
      );

      if (!response.ok) {
        if (response.status === 401) {
          // Token expired, try to reinitialize
          await MicrosoftApiService.initialize();
          return this.updateDrawing(drawing);
        }
        throw new Error(`Failed to update drawing: ${response.statusText}`);
      }

      XLogger.log("Drawing updated in Microsoft OneDrive successfully");
    } catch (error) {
      XLogger.error("Error updating drawing in Microsoft OneDrive", error);
      throw error;
    }
  }

  static async deleteDrawing(drawingName: string): Promise<void> {
    if (!MicrosoftApiService.accessToken) {
      await MicrosoftApiService.initialize();
    }

    try {
      const response = await fetch(
        `${MicrosoftApiService.API_BASE_URL}/me/drive/root:/ExcaliSave/${drawingName}.json`,
        {
          method: "DELETE",
          headers: {
            Authorization: `Bearer ${MicrosoftApiService.accessToken}`,
          },
        }
      );

      if (!response.ok) {
        if (response.status === 401) {
          // Token expired, try to reinitialize
          await MicrosoftApiService.initialize();
          return this.deleteDrawing(drawingName);
        }
        throw new Error(`Failed to delete drawing: ${response.statusText}`);
      }

      XLogger.log("Drawing deleted from Microsoft OneDrive successfully");
    } catch (error) {
      XLogger.error("Error deleting drawing from Microsoft OneDrive", error);
      throw error;
    }
  }

  public static async generatePKCE(): Promise<{
    codeVerifier: string;
    codeChallenge: string;
  }> {
    // Generate a random string for code verifier
    const array = new Uint8Array(32);
    crypto.getRandomValues(array);
    const codeVerifier = Array.from(array, (b) =>
      b.toString(16).padStart(2, "0")
    ).join("");

    // Generate code challenge
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    const digest = await crypto.subtle.digest("SHA-256", data);
    const digestArray = new Uint8Array(digest);
    const codeChallenge = btoa(
      String.fromCharCode.apply(null, Array.from(digestArray))
    )
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");

    return { codeVerifier, codeChallenge };
  }

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
