import { XLogger } from "./logger";
import { browser } from "webextension-polyfill-ts";
import { MicrosoftFetchService } from "./microsoft-fetch.service";

export class MicrosoftAuthService {
  private static instance: MicrosoftAuthService;
  public static accessToken: string | null = null;
  public static initialized = false;
  public static readonly TENANT_ID = "common";
  public static readonly CLIENT_ID = "YOUR_CLIENT_ID";
  public static readonly REDIRECT_URI = "https://excalidraw.com/auth-callback";
  public static readonly SCOPES = ["Files.ReadWrite", "User.Read"];

  private constructor() {}

  public static getInstance(): MicrosoftAuthService {
    if (!MicrosoftAuthService.instance) {
      MicrosoftAuthService.instance = new MicrosoftAuthService();
    }
    return MicrosoftAuthService.instance;
  }

  public static async initialize(): Promise<void> {
    if (MicrosoftAuthService.initialized) {
      return;
    }

    try {
      const result = await browser.storage.local.get("microsoftAccessToken");
      if (result.microsoftAccessToken) {
        MicrosoftAuthService.accessToken = result.microsoftAccessToken;
        MicrosoftAuthService.initialized = true;
      } else {
        MicrosoftAuthService.accessToken = null;
        MicrosoftAuthService.initialized = false;
      }
    } catch (error) {
      XLogger.error("Error initializing Microsoft Auth service", error);
      MicrosoftAuthService.accessToken = null;
      MicrosoftAuthService.initialized = false;
    }
  }

  public async saveAccessToken(token: string): Promise<void> {
    try {
      await browser.storage.local.set({ microsoftAccessToken: token });
      MicrosoftAuthService.accessToken = token;
      MicrosoftAuthService.initialized = true;
    } catch (error) {
      XLogger.error("Error saving Microsoft access token", error);
      throw error;
    }
  }

  public async removeAccessToken(): Promise<void> {
    try {
      await browser.storage.local.remove("microsoftAccessToken");
      MicrosoftAuthService.accessToken = null;
      MicrosoftAuthService.initialized = false;
    } catch (error) {
      XLogger.error("Error removing Microsoft access token", error);
      throw error;
    }
  }

  public static isAuthenticated(): boolean {
    return (
      MicrosoftAuthService.initialized &&
      MicrosoftAuthService.accessToken !== null
    );
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
    return MicrosoftFetchService.exchangeCodeForToken(
      code,
      redirectUri,
      codeVerifier,
      clientId,
      tenantId
    );
  }

  public static async ensureInitialized(): Promise<void> {
    if (!MicrosoftAuthService.accessToken) {
      await MicrosoftAuthService.initialize();
    }
  }
}
