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
} from "@radix-ui/themes";
import React, { useEffect, useState } from "react";
import { browser } from "webextension-polyfill-ts";
import { ImpExp } from "../components/ImpExp/ImpExp.component";
import "./Options.styles.scss";
import { XLogger } from "../lib/logger";
import { MicrosoftApiService } from "../lib/microsoft-api.service";
import { MicrosoftAuthService } from "../lib/microsoft-auth.service";
import { Folder } from "../interfaces/folder.interface";

export function Options() {
  const [accessToken, setAccessToken] = useState<string>("");
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [tenantId, setTenantId] = useState<string>("organizations");
  const [appId, setAppId] = useState<string>("YOUR_CLIENT_ID");
  const [codeVerifier, setCodeVerifier] = useState<string>("");

  useEffect(() => {
    const init = async () => {
      try {
        await MicrosoftApiService.initialize();
        setIsAuthenticated(MicrosoftAuthService.isAuthenticated());

        // Initialize form fields with values from the service
        setTenantId(MicrosoftAuthService.TENANT_ID);
        setAppId(MicrosoftAuthService.CLIENT_ID);

        // Check for access token in URL fragment
        const hash = window.location.hash;
        if (hash) {
          const params = new URLSearchParams(hash.substring(1));
          const token = params.get("access_token");
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
      } catch (error) {
        XLogger.error("Error initializing Microsoft API", error);
      } finally {
        setIsLoading(false);
      }
    };
    init();
  }, []);

  const handleOpenAuthPage = async () => {
    const redirectUri = "https://excalidraw.com/auth-callback";
    const { codeVerifier, codeChallenge } =
      await MicrosoftAuthService.generatePKCE();
    setCodeVerifier(codeVerifier);

    const authUrl =
      `https://login.microsoftonline.com/${MicrosoftAuthService.TENANT_ID}/oauth2/v2.0/authorize` +
      `?client_id=${MicrosoftAuthService.CLIENT_ID}` +
      `&response_type=code` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&scope=${MicrosoftAuthService.SCOPES.join(" ")}` +
      `&response_mode=query` +
      `&code_challenge=${codeChallenge}` +
      `&code_challenge_method=S256`;

    // Open auth page in a new window
     window.open(authUrl, "_blank", "width=600,height=600");

    // Show instructions to the user
    alert(
      "After logging in, you will be redirected to a page that doesn't exist.\n" +
        "Please copy the authorization code from the URL in the address bar and paste it below."
    );
  };

  const handleSaveToken = async (code: string) => {
    try {
      const redirectUri = "https://excalidraw.com/auth-callback";
      const token = await MicrosoftAuthService.exchangeCodeForToken(
        code,
        redirectUri,
        codeVerifier,
        MicrosoftAuthService.CLIENT_ID,
        MicrosoftAuthService.TENANT_ID
      );
      await MicrosoftAuthService.getInstance().saveAccessToken(token);
      setIsAuthenticated(true);
      setAccessToken("");
      setCodeVerifier("");

      // Create sync folder if it doesn't exist
      const folders = await browser.storage.local.get("folders");
      const syncFolder = folders.folders?.find(
        (f: any) => f.name === MicrosoftApiService.SYNC_FOLDER_NAME
      );

      if (!syncFolder) {
        const newFolder: Folder = {
          id: `folder:${Math.random().toString(36).substr(2, 9)}`,
          name: MicrosoftApiService.SYNC_FOLDER_NAME,
          drawingIds: [],
        };

        const newFolders = [...(folders.folders || []), newFolder];
        await browser.storage.local.set({ folders: newFolders });
      }
    } catch (error) {
      XLogger.error("Error saving access token", error);
    }
  };

  const handleLogout = async () => {
    try {
      await MicrosoftAuthService.getInstance().removeAccessToken();
      setIsAuthenticated(false);
    } catch (error) {
      XLogger.error("Error removing access token", error);
    }
  };

  const handleTenantIdChange = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const newTenantId = e.target.value;
    setTenantId(newTenantId);
    await MicrosoftAuthService.setTenantId(newTenantId);
  };

  const handleAppIdChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const newAppId = e.target.value;
    setAppId(newAppId);
    await MicrosoftAuthService.setClientId(newAppId);
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
              Microsoft OneDrive Integration:
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
                      Open Microsoft Login Page
                    </Button>
                  </li>
                  <li>
                    <Text mb="2">
                      After logging in, copy the authorization code from the URL
                      and paste it below:
                    </Text>
                    <TextField.Root>
                      <TextField.Input
                        value={accessToken}
                        onChange={(e) => setAccessToken(e.target.value)}
                        placeholder="Paste authorization code here"
                      />
                    </TextField.Root>
                  </li>
                  <li>
                    <Button onClick={() => handleSaveToken(accessToken)} mt="2">
                      Save Token
                    </Button>
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
    </Theme>
  );
}
