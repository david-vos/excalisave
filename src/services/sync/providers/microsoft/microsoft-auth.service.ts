import { browser } from "webextension-polyfill-ts";
import { XLogger } from "../../../../lib/logger";

/**
 * Service for handling Microsoft authentication
 */
export class MicrosoftAuthService {
  public static readonly TENANT_ID = "organizations";
  public static readonly CLIENT_ID = "YOUR_CLIENT_ID";
  public static readonly SCOPES = [
    "Files.ReadWrite",
    "User.Read",
    "offline_access",
  ];
  public static readonly REDIRECT_URI = "https://excalidraw.com/";

  private static instance: MicrosoftAuthService;
  private accessToken: string | null = null;
  private clientId: string | null = null;
  private tenantId: string | null = null;

  private constructor() {
    this.loadStoredValues();
  }

  /**
   * Set the tenant ID
   */
  public async setTenantId(tenantId: string): Promise<void> {
    this.tenantId = tenantId;
    await browser.storage.local.set({ microsoftTenantId: tenantId });
  }

  /**
   * Set the client ID
   */
  public async setClientId(clientId: string): Promise<void> {
    this.clientId = clientId;
    await browser.storage.local.set({ microsoftClientId: clientId });
  }

  private async loadStoredValues(): Promise<void> {
    try {
      const result = await browser.storage.local.get([
        "microsoftClientId",
        "microsoftTenantId",
      ]);

      if (result.microsoftClientId) {
        this.clientId = result.microsoftClientId;
      }

      if (result.microsoftTenantId) {
        this.tenantId = result.microsoftTenantId;
      }
    } catch (error) {
      XLogger.error(
        `Error loading stored Microsoft auth values: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  /**
   * Get the singleton instance of the MicrosoftAuthService
   */
  public static getInstance(): MicrosoftAuthService {
    if (!MicrosoftAuthService.instance) {
      MicrosoftAuthService.instance = new MicrosoftAuthService();
    }
    return MicrosoftAuthService.instance;
  }

  /**
   * Save the access token
   */
  public async saveAccessToken(token: string): Promise<void> {
    this.accessToken = token;
    await browser.storage.local.set({ microsoftAccessToken: token });
  }

  /**
   * Remove the access token
   */
  public async removeAccessToken(): Promise<void> {
    this.accessToken = null;
    await browser.storage.local.remove("microsoftAccessToken");
  }

  /**
   * Check if the user is authenticated
   */
  public async isAuthenticated(): Promise<boolean> {
    if (this.accessToken) {
      return true;
    }

    const result = await browser.storage.local.get("microsoftAccessToken");
    if (result.microsoftAccessToken) {
      this.accessToken = result.microsoftAccessToken;
      return true;
    }

    return false;
  }

  /**
   * Get the access token
   */
  public async getAccessToken(): Promise<string | null> {
    if (this.accessToken) {
      return this.accessToken;
    }

    const result = await browser.storage.local.get("microsoftAccessToken");
    if (result.microsoftAccessToken) {
      this.accessToken = result.microsoftAccessToken;
      return this.accessToken;
    }

    return null;
  }

  /**
   * Get the client ID
   */
  public async getClientId(): Promise<string> {
    if (this.clientId) {
      return this.clientId;
    }

    await this.loadStoredValues();
    return this.clientId || MicrosoftAuthService.CLIENT_ID;
  }

  /**
   * Get the tenant ID
   */
  public async getTenantId(): Promise<string> {
    if (this.tenantId) {
      return this.tenantId;
    }

    await this.loadStoredValues();
    return this.tenantId || MicrosoftAuthService.TENANT_ID;
  }

  /**
   * Generate a PKCE code verifier and challenge
   */
  public async generatePKCE(): Promise<{
    codeVerifier: string;
    codeChallenge: string;
  }> {
    // Generate a random string for the code verifier (43-128 characters)
    const codeVerifier = this.generateRandomString(64);

    // Create an SHA-256 hash of the code verifier
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    const hash = await crypto.subtle.digest("SHA-256", data);
    const hashArray = Array.from(new Uint8Array(hash));

    // Base64URL encode the hash
    const base64Hash = btoa(String.fromCharCode.apply(null, hashArray))
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");

    return { codeVerifier, codeChallenge: base64Hash };
  }

  private generateRandomString(length: number): string {
    const possible =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
    let text = "";
    for (let i = 0; i < length; i++) {
      text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
  }

  /**
   * Exchange an authorization code for an access token
   */
  public async exchangeCodeForToken(
    code: string,
    codeVerifier: string,
    clientId: string,
    tenantId: string
  ): Promise<string> {
    try {
      XLogger.info("Exchanging code for token with:", {
        code: code.substring(0, 10) + "...",
        codeVerifier: codeVerifier.substring(0, 10) + "...",
        clientId,
        tenantId,
        redirectUri: MicrosoftAuthService.REDIRECT_URI,
      });

      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const params = new URLSearchParams({
        client_id: clientId,
        scope: MicrosoftAuthService.SCOPES.join(" "),
        code,
        redirect_uri: MicrosoftAuthService.REDIRECT_URI,
        grant_type: "authorization_code",
        code_verifier: codeVerifier,
      });

      XLogger.info("Token endpoint:", tokenEndpoint);
      XLogger.info("Request params:", params.toString());

      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params.toString(),
      });

      XLogger.info("Token response status:", response.status);

      if (!response.ok) {
        const errorText = await response.text();
        XLogger.error("Token exchange failed:", {
          status: response.status,
          statusText: response.statusText,
          errorText,
        });
        throw new Error(
          `Token exchange failed: ${response.statusText} - ${errorText}`
        );
      }

      const data = await response.json();
      XLogger.info("Token response data:", {
        tokenType: data.token_type,
        expiresIn: data.expires_in,
        scope: data.scope,
        hasAccessToken: !!data.access_token,
      });

      if (!data.access_token) {
        throw new Error("No access token in response");
      }

      return data.access_token;
    } catch (error) {
      XLogger.error(
        `Error exchanging code for token: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      throw error;
    }
  }
}
