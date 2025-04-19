import {
  Avatar,
  Box,
  Container,
  Flex,
  Heading,
  Text,
  Theme,
  Button,
  TextField,
  AlertDialog,
  AlertDialogContent,
  AlertDialogTitle,
  AlertDialogDescription,
  AlertDialogAction,
} from "@radix-ui/themes";
import React, { useEffect, useState } from "react";
import { browser } from "webextension-polyfill-ts";
import { ImpExp } from "../components/ImpExp/ImpExp.component";
import "./Options.styles.scss";
import { XLogger } from "../lib/logger";
import { Folder } from "../interfaces/folder.interface";
import { SyncService } from "../services/sync";
import {
  MicrosoftProvider,
  MicrosoftAuthService,
} from "../services/sync/providers/microsoft";

export function Options() {
  const [accessToken, setAccessToken] = useState<string>("");
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [tenantId, setTenantId] = useState<string>("organizations");
  const [appId, setAppId] = useState<string>("YOUR_CLIENT_ID");
  const [codeVerifier, setCodeVerifier] = useState<string>("");
  const [isAuthInProgress, setIsAuthInProgress] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [showErrorDialog, setShowErrorDialog] = useState<boolean>(false);
  const [debugInfo, setDebugInfo] = useState<string>("");

  // Function to store code verifier in browser storage
  const storeCodeVerifier = async (verifier: string) => {
    try {
      await browser.storage.local.set({ microsoftCodeVerifier: verifier });
      XLogger.info("Stored code verifier in browser storage");
    } catch (error) {
      XLogger.error("Error storing code verifier:", error);
    }
  };

  // Function to retrieve code verifier from browser storage
  const retrieveCodeVerifier = async (): Promise<string | null> => {
    try {
      const storedCodeVerifier = await browser.storage.local.get(
        "microsoftCodeVerifier"
      );
      if (storedCodeVerifier.microsoftCodeVerifier) {
        XLogger.info("Retrieved code verifier from browser storage");
        return storedCodeVerifier.microsoftCodeVerifier;
      }
      return null;
    } catch (error) {
      XLogger.error("Error retrieving code verifier:", error);
      return null;
    }
  };

  // Function to clear code verifier from browser storage
  const clearCodeVerifier = async () => {
    try {
      await browser.storage.local.remove("microsoftCodeVerifier");
      XLogger.info("Cleared code verifier from browser storage");
    } catch (error) {
      XLogger.error("Error clearing code verifier:", error);
    }
  };

  useEffect(() => {
    const init = async () => {
      try {
        XLogger.info("Initializing Options component...");
        // Initialize sync service with Microsoft provider
        const syncService = SyncService.getInstance();
        syncService.setProvider(MicrosoftProvider.getInstance());
        await syncService.initialize();
        XLogger.info("Sync service initialized");

        const authStatus = await syncService.isAuthenticated();
        XLogger.info("Authentication status:", authStatus);
        setIsAuthenticated(authStatus);

        // Initialize form fields with values from the service
        const authService = MicrosoftAuthService.getInstance();
        const storedTenantId = await authService.getTenantId();
        const storedAppId = await authService.getClientId();
        XLogger.info("Stored tenant ID:", storedTenantId);
        XLogger.info("Stored app ID:", storedAppId);
        setTenantId(storedTenantId);
        setAppId(storedAppId);

        // Check for access token in URL fragment
        const hash = window.location.hash;
        XLogger.info("URL hash:", hash);
        if (hash) {
          const params = new URLSearchParams(hash.substring(1));
          const token = params.get("access_token");
          XLogger.info("Access token in URL:", token ? "Found" : "Not found");
          if (token) {
            await handleSaveToken(token);
            // Clear the URL fragment
            window.history.replaceState(
              {},
              document.title,
              window.location.pathname
            );
          }
        }

        // Check for authorization code in URL
        const url = new URL(window.location.href);
        XLogger.info("Current URL:", url.href);
        if (url.searchParams.has("code")) {
          const code = url.searchParams.get("code");
          XLogger.info(
            "Authorization code in URL:",
            code ? "Found" : "Not found"
          );
          if (code) {
            // Get the stored code verifier from browser storage
            const storedCodeVerifier = await retrieveCodeVerifier();
            if (storedCodeVerifier) {
              XLogger.info("Found stored code verifier");
              setCodeVerifier(storedCodeVerifier);
              await handleSaveToken(code);
              // Clear the URL parameters
              window.history.replaceState(
                {},
                document.title,
                window.location.pathname
              );
            } else {
              XLogger.error("No stored code verifier found");
              setErrorMessage(
                "Authentication failed: No code verifier found. Please try again."
              );
              setShowErrorDialog(true);
            }
          }
        }
      } catch (error) {
        XLogger.error("Error initializing Microsoft API:", error);
        setErrorMessage(
          `Error initializing: ${
            error instanceof Error ? error.message : String(error)
          }`
        );
        setShowErrorDialog(true);
      } finally {
        setIsLoading(false);
      }
    };
    init();
  }, []);

  // Check for authorization code in URL when component mounts or auth is in progress
  useEffect(() => {
    if (isAuthInProgress) {
      XLogger.info("Auth in progress, checking for authorization code...");
      const checkForAuthCode = () => {
        const url = new URL(window.location.href);
        XLogger.info("Checking URL for auth code:", url.href);
        // Check if the URL contains the authorization code, regardless of the domain
        if (url.searchParams.has("code")) {
          const code = url.searchParams.get("code");
          XLogger.info("Authorization code found:", code);
          if (code) {
            handleSaveToken(code);
            setIsAuthInProgress(false);
            // Clear the URL parameters
            window.history.replaceState(
              {},
              document.title,
              window.location.pathname
            );
          }
        }
      };

      // Check immediately
      checkForAuthCode();

      // Set up an interval to check periodically
      const interval = setInterval(checkForAuthCode, 1000);

      // Clean up interval on component unmount or when auth is no longer in progress
      return () => clearInterval(interval);
    }

    // Return a no-op cleanup function when not in auth progress
    return () => {};
  }, [isAuthInProgress]);

  const handleOpenAuthPage = async () => {
    try {
      XLogger.info("Opening auth page...");
      setIsLoading(true);
      setIsAuthInProgress(true);
      setErrorMessage("");
      setDebugInfo("");

      // Generate PKCE code verifier and challenge
      const { codeVerifier, codeChallenge } =
        await MicrosoftAuthService.getInstance().generatePKCE();
      XLogger.info("Generated PKCE code verifier:", codeVerifier);
      XLogger.info("Generated PKCE code challenge:", codeChallenge);
      setCodeVerifier(codeVerifier);

      // Store the code verifier in browser storage for persistence
      await storeCodeVerifier(codeVerifier);

      // Use the values from the UI state
      XLogger.info("Using client ID from UI:", appId);
      XLogger.info("Using tenant ID from UI:", tenantId);

      // Construct the authorization URL
      const authUrl =
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
        `?client_id=${appId}` +
        `&response_type=code` +
        `&redirect_uri=${encodeURIComponent(
          MicrosoftAuthService.REDIRECT_URI
        )}` +
        `&scope=${MicrosoftAuthService.SCOPES.join(" ")}` +
        `&response_mode=query` +
        `&code_challenge=${codeChallenge}` +
        `&code_challenge_method=S256`;

      XLogger.info("Authorization URL:", authUrl);
      setDebugInfo(
        `Auth URL: ${authUrl}\nRedirect URI: ${MicrosoftAuthService.REDIRECT_URI}`
      );

      // Open the auth page in a new tab
      const authTab = await browser.tabs.create({
        url: authUrl,
        active: true,
      });

      XLogger.info("Opened auth tab with ID:", authTab.id);

      // Set up a listener for the redirect
      const tabListener = (tabId: number, changeInfo: any) => {
        if (
          tabId === authTab.id &&
          changeInfo.url &&
          changeInfo.url.includes(MicrosoftAuthService.REDIRECT_URI)
        ) {
          XLogger.info("Detected redirect to:", changeInfo.url);

          // Extract the authorization code from the redirect URL
          const url = new URL(changeInfo.url);
          const code = url.searchParams.get("code");

          if (code) {
            XLogger.info("Authorization code found:", code);

            // Remove the listener
            browser.tabs.onUpdated.removeListener(tabListener);

            // Close the auth tab
            browser.tabs.remove(tabId);

            // Handle the authorization code
            handleSaveToken(code);
          } else {
            XLogger.error("No authorization code found in redirect URL");
            setErrorMessage(
              "Authentication failed: No authorization code found"
            );
            setShowErrorDialog(true);
            setIsLoading(false);
            setIsAuthInProgress(false);
          }
        }
      };

      // Add the listener
      browser.tabs.onUpdated.addListener(tabListener);

      // Set a timeout to clean up if the user doesn't complete the flow
      setTimeout(() => {
        browser.tabs.onUpdated.removeListener(tabListener);
        if (isAuthInProgress) {
          XLogger.warn("Authentication timed out");
          setErrorMessage("Authentication timed out. Please try again.");
          setShowErrorDialog(true);
          setIsLoading(false);
          setIsAuthInProgress(false);
        }
      }, 300000); // 5 minutes timeout
    } catch (error) {
      XLogger.error("Error during authentication:", error);
      setIsLoading(false);
      setIsAuthInProgress(false);
      setErrorMessage(
        `Error during authentication: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      setShowErrorDialog(true);
    }
  };

  const handleSaveToken = async (code: string) => {
    try {
      XLogger.info("Saving token with code:", code);
      setIsLoading(true);
      setErrorMessage("");

      // Use the values from the UI state
      XLogger.info("Using client ID from UI for token exchange:", appId);
      XLogger.info("Using tenant ID from UI for token exchange:", tenantId);
      XLogger.info("Using code verifier:", codeVerifier);

      // If codeVerifier is not available in state, try to get it from storage
      let verifierToUse = codeVerifier;
      if (!verifierToUse) {
        verifierToUse = await retrieveCodeVerifier();
        if (!verifierToUse) {
          throw new Error("No code verifier available for token exchange");
        }
      }

      const token =
        await MicrosoftAuthService.getInstance().exchangeCodeForToken(
          code,
          verifierToUse,
          appId,
          tenantId
        );
      XLogger.info("Token received:", token ? "Success" : "Failed");
      await MicrosoftAuthService.getInstance().saveAccessToken(token);
      setIsAuthenticated(true);
      setAccessToken(token);

      // Clear the stored code verifier after successful token exchange
      await clearCodeVerifier();

      // Create sync folder if it doesn't exist
      const folders = await browser.storage.local.get("folders");
      XLogger.info("Current folders:", folders);
      const syncFolder = folders.folders?.find(
        (f: any) => f.name === SyncService.SYNC_FOLDER_NAME
      );
      XLogger.info("Sync folder exists:", !!syncFolder);

      if (!syncFolder) {
        XLogger.info("Creating new sync folder");
        const newFolder = {
          id: `folder:${Math.random().toString(36).substr(2, 9)}`,
          name: SyncService.SYNC_FOLDER_NAME,
          drawingIds: [] as string[],
        };

        const newFolders = [...(folders.folders || []), newFolder];
        await browser.storage.local.set({ folders: newFolders });
        XLogger.info("New sync folder created");
      }
    } catch (error) {
      XLogger.error("Error saving token:", error);
      setErrorMessage(
        `Error saving token: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      setShowErrorDialog(true);
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = async () => {
    try {
      XLogger.info("Logging out...");
      setIsLoading(true);
      await MicrosoftAuthService.getInstance().removeAccessToken();
      setIsAuthenticated(false);
      setAccessToken("");
      XLogger.info("Logout successful");
    } catch (error) {
      XLogger.error("Error logging out:", error);
      setErrorMessage(
        `Error logging out: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
      setShowErrorDialog(true);
    } finally {
      setIsLoading(false);
    }
  };

  const handleTenantIdChange = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const newTenantId = e.target.value;
    XLogger.info("Tenant ID changed to:", newTenantId);
    setTenantId(newTenantId);
    await MicrosoftAuthService.getInstance().setTenantId(newTenantId);
  };

  const handleAppIdChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const newAppId = e.target.value;
    XLogger.info("App ID changed to:", newAppId);
    setAppId(newAppId);
    await MicrosoftAuthService.getInstance().setClientId(newAppId);
  };

  const closeErrorDialog = () => {
    setShowErrorDialog(false);
  };

  if (isLoading) {
    return <Text>Loading...</Text>;
  }

  return (
    <Theme
      accentColor="iris"
      style={{
        height: "100%",
      }}
    >
      <Box
        style={{
          background: "var(--gray-a2)",
          borderRadius: "var(--radius-3)",
          width: "100vw",
          height: "100vh",
        }}
      >
        <Container size="2">
          <Flex gap="3" px="1" py="9" justify={"start"} align={"center"}>
            <Box>
              <Avatar
                size={"5"}
                src={browser.runtime.getURL("assets/icons/128.png")}
                fallback={""}
              />
            </Box>
            <Box>
              <Heading as="h1" size="7" style={{ paddingBottom: "4px" }}>
                ExcaliSave Settings
              </Heading>
              <Text size={"2"} as="p" style={{ lineHeight: 1.1 }}>
                Customize how the ExcaliSave extension works in your browser.
                <br />
                These settings are specific to this browser profile.
              </Text>
            </Box>
          </Flex>
          <Box px="4">
            <Heading as="h3" size={"5"} style={{ paddingBottom: "10px" }}>
              Import/Export:
            </Heading>
            <Text size={"2"} as="p" style={{ lineHeight: 1.1 }}>
              Import or export your data to or from ExcaliSave.
            </Text>
            <br />
            <ImpExp />
          </Box>
          <Box px="4">
            <Heading as="h3" size={"5"} style={{ paddingBottom: "10px" }}>
              Cloud Sync Integration:
            </Heading>
            {!isAuthenticated ? (
              <Box>
                <Text size="2" as="p" style={{ marginBottom: "10px" }}>
                  To connect your Microsoft account:
                </Text>
                <ol style={{ marginBottom: "10px" }}>
                  <li>
                    <Box mb="2">
                      <Text mb="1">Tenant ID:</Text>
                      <TextField.Root>
                        <TextField.Input
                          value={tenantId}
                          onChange={handleTenantIdChange}
                          placeholder="Enter your Microsoft Tenant ID"
                        />
                      </TextField.Root>
                    </Box>
                    <Box mb="2">
                      <Text mb="1">App ID:</Text>
                      <TextField.Root>
                        <TextField.Input
                          value={appId}
                          onChange={handleAppIdChange}
                          placeholder="Enter your Microsoft App ID"
                        />
                      </TextField.Root>
                    </Box>
                    <Button onClick={handleOpenAuthPage} mb="2">
                      Connect to Microsoft
                    </Button>
                    <Text size="1" color="gray">
                      You will be redirected to Microsoft for authentication.
                    </Text>
                    {debugInfo && (
                      <Box
                        mt="2"
                        p="2"
                        style={{ background: "#f5f5f5", borderRadius: "4px" }}
                      >
                        <Text
                          size="1"
                          style={{
                            whiteSpace: "pre-wrap",
                            fontFamily: "monospace",
                          }}
                        >
                          {debugInfo}
                        </Text>
                      </Box>
                    )}
                  </li>
                </ol>
              </Box>
            ) : (
              <Box>
                <Text mb="2">You are connected to Microsoft OneDrive.</Text>
                <Button onClick={handleLogout} variant="soft" color="red">
                  Disconnect
                </Button>
              </Box>
            )}
          </Box>
        </Container>
      </Box>

      {/* Error Dialog */}
      <AlertDialog.Root open={showErrorDialog}>
        <AlertDialogContent>
          <AlertDialogTitle>Error</AlertDialogTitle>
          <Box>
            <Text>{errorMessage}</Text>
          </Box>
          <AlertDialogAction onClick={closeErrorDialog}>OK</AlertDialogAction>
        </AlertDialogContent>
      </AlertDialog.Root>
    </Theme>
  );
}
