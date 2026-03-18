import { Tabs, browser } from "webextension-polyfill-ts";
import { XLogger } from "../logger";

export class TabUtils {
  static async getActiveTab(): Promise<Tabs.Tab | undefined> {
    const tabs = await browser.tabs.query({
      active: true,
      currentWindow: true,
    });

    if (tabs.length === 0) {
      XLogger.warn("No active tab found");
      return undefined;
    }

    return tabs[0];
  }
}
