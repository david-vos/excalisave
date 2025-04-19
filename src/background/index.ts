import { browser } from "webextension-polyfill-ts";
import {
  CleanupFilesMessage,
  MessageType,
  SaveDrawingMessage,
  SaveNewDrawingMessage,
} from "../constants/message.types";
import { IDrawing } from "../interfaces/drawing.interface";
import { XLogger } from "../lib/logger";
import { TabUtils } from "../lib/utils/tab.utils";
import { RandomUtils } from "../lib/utils/random.utils";
import { SyncService } from "../services/sync";
import { MicrosoftProvider } from "../services/sync/providers/microsoft";

// Initialize sync service with Microsoft provider
const syncService = SyncService.getInstance();
syncService.setProvider(MicrosoftProvider.getInstance());

browser.runtime.onInstalled.addListener(async () => {
  XLogger.info("Extension installed or updated");

  for (const cs of (browser.runtime.getManifest() as any).content_scripts) {
    for (const tab of await browser.tabs.query({ url: cs.matches })) {
      browser.scripting.executeScript({
        target: { tabId: tab.id },
        files: cs.js,
      });
    }
  }

  // Initialize sync service
  try {
    await syncService.initialize();
  } catch (error) {
    XLogger.error(
      `Error initializing sync service: ${
        error instanceof Error ? error.message : String(error)
      }`
    );
  }
});

browser.runtime.onMessage.addListener(
  async (
    message:
      | SaveDrawingMessage
      | SaveNewDrawingMessage
      | CleanupFilesMessage
      | any,
    _sender: any
  ) => {
    try {
      XLogger.info(
        `Message received in background: ${message.type || "unknown"}`
      );
      if (!message || !message.type) return;

      switch (message.type) {
        case MessageType.SAVE_NEW_DRAWING:
          const newDrawing = {
            id: message.payload.id,
            name: message.payload.name,
            createdAt: new Date().toISOString(),
            imageBase64: message.payload.imageBase64,
            viewBackgroundColor: message.payload.viewBackgroundColor,
            data: {
              excalidraw: message.payload.excalidraw,
              excalidrawState: message.payload.excalidrawState,
              versionFiles: message.payload.versionFiles,
              versionDataState: message.payload.versionDataState,
            },
          };

          // Save to browser storage
          await browser.storage.local.set({
            [message.payload.id]: newDrawing,
          });

          // Save to cloud
          try {
            await syncService.saveDrawing(newDrawing);
          } catch (error) {
            XLogger.error(
              `Failed to save to cloud: ${
                error instanceof Error ? error.message : String(error)
              }`
            );
            // Continue execution even if cloud save fails
          }
          break;

        case MessageType.SAVE_DRAWING:
          const existentDrawing = (
            await browser.storage.local.get(message.payload.id)
          )[message.payload.id] as IDrawing;

          if (!existentDrawing) {
            XLogger.error(`No drawing found with id: ${message.payload.id}`);
            return;
          }

          const updatedDrawing: IDrawing = {
            ...existentDrawing,
            name: message.payload.name || existentDrawing.name,
            imageBase64:
              message.payload.imageBase64 || existentDrawing.imageBase64,
            viewBackgroundColor:
              message.payload.viewBackgroundColor ||
              existentDrawing.viewBackgroundColor,
            data: {
              excalidraw: message.payload.excalidraw,
              excalidrawState: message.payload.excalidrawState,
              versionFiles: message.payload.versionFiles,
              versionDataState: message.payload.versionDataState,
            },
          };

          // Update in browser storage
          await browser.storage.local.set({
            [message.payload.id]: updatedDrawing,
          });

          // Update in cloud
          try {
            await syncService.updateDrawing(updatedDrawing);
          } catch (error) {
            XLogger.error(
              `Failed to update in cloud: ${
                error instanceof Error ? error.message : String(error)
              }`
            );
            // Continue execution even if cloud update fails
          }
          break;

        case MessageType.CLEANUP_FILES:
          XLogger.info("Cleaning up files");

          const drawings = Object.values(
            await browser.storage.local.get()
          ).filter((o) => o?.id?.startsWith?.("drawing:"));

          const imagesUsed = drawings
            .map((drawing) => {
              return JSON.parse(drawing.data.excalidraw).filter(
                (item: any) => item.type === "image"
              );
            })
            .flat()
            .map<string>((item) => item?.fileId);

          const uniqueImagesUsed = Array.from(new Set(imagesUsed));

          XLogger.info(`Used fileIds: ${uniqueImagesUsed.length}`);

          // This workaround is to pass params to script, it's ugly but it works
          await browser.scripting.executeScript({
            target: {
              tabId: message.payload.tabId,
            },
            func: (fileIds: string[], executionTimestamp: number) => {
              window.__SCRIPT_PARAMS__ = { fileIds, executionTimestamp };
            },
            args: [uniqueImagesUsed, message.payload.executionTimestamp],
          });

          await browser.scripting.executeScript({
            target: { tabId: message.payload.tabId },
            files: ["./js/execute-scripts/delete-unused-files.bundle.js"],
          });

          break;

        case "MessageAutoSave":
          const name = message.payload.name;
          const setCurrent = message.payload.setCurrent;
          XLogger.info(`Saving new drawing: ${name}`);
          const activeTab = await TabUtils.getActiveTab();

          if (!activeTab) {
            XLogger.warn("No active tab found");
            return;
          }

          const id = `drawing:${RandomUtils.generateRandomId()}`;

          // This workaround is to pass params to script, it's ugly but it works
          await browser.scripting.executeScript({
            target: { tabId: activeTab.id },
            func: (id, name, setCurrent) => {
              window.__SCRIPT_PARAMS__ = { id, name, setCurrent };
            },
            args: [id, name, setCurrent],
          });

          await browser.scripting.executeScript({
            target: { tabId: activeTab.id },
            files: ["./js/execute-scripts/sendDrawingDataToSave.bundle.js"],
          });
          break;
        default:
          break;
      }
    } catch (error) {
      XLogger.error(
        `Error handling message: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }
);
