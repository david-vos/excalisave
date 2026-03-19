import React, { useState, useEffect, useRef, useCallback } from "react";
import "./DrawingTitle.styles.scss";
import { useLocalStorageString } from "../../hooks/useLocalStorageString.hook";
import {
  DRAWING_ID_KEY_LS,
  DRAWING_TITLE_KEY_LS,
} from "../../../lib/constants";
import { browser } from "webextension-polyfill-ts";
import { MessageType, SaveDrawingMessage } from "../../../constants/message.types";
import { getDrawingDataState } from "../../content-script.utils";
import { XLogger } from "../../../lib/logger";
import { As } from "../../../lib/types.utils";

type DrawingListItem = {
  id: string;
  name: string;
  createdAt?: string;
  roomUrl?: string;
};

function getDefaultDrawingName(): string {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `Drawing - ${yyyy}-${mm}-${dd}`;
}

export function DrawingTitle() {
  const title = useLocalStorageString(DRAWING_TITLE_KEY_LS, "");
  const currentDrawingId = useLocalStorageString(DRAWING_ID_KEY_LS, "");
  const [isOpen, setIsOpen] = useState(false);
  const [drawings, setDrawings] = useState<DrawingListItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [newName, setNewName] = useState("");
  const [search, setSearch] = useState("");
  const [searchResults, setSearchResults] = useState<DrawingListItem[] | null>(null);
  const [showTooltip, setShowTooltip] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const isUnsaved = !currentDrawingId;

  useEffect(() => {
    if (!isOpen) return undefined;

    const handleClickOutside = (e: MouseEvent) => {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(e.target as Node)
      ) {
        setIsOpen(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return undefined;
    setNewName(getDefaultDrawingName());
    setSearch("");

    setLoading(true);
    browser.runtime
      .sendMessage({ type: MessageType.GET_ALL_DRAWINGS })
      .then((response: any) => {
        if (response?.success) {
          const sorted = (response.drawings || []).sort(
            (a: DrawingListItem, b: DrawingListItem) => {
              if (!a.createdAt) return 1;
              if (!b.createdAt) return -1;
              return b.createdAt.localeCompare(a.createdAt);
            }
          );
          setDrawings(sorted);
        }
      })
      .catch((err) => XLogger.error("Failed to fetch drawings", err))
      .finally(() => setLoading(false));
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return undefined;

    const query = search.trim();
    if (!query) {
      setSearchResults(null);
      return undefined;
    }

    const timer = setTimeout(() => {
      browser.runtime
        .sendMessage({
          type: MessageType.SEARCH_DRAWINGS,
          payload: { query },
        })
        .then((response: any) => {
          if (response?.success) {
            setSearchResults(response.drawings || []);
          }
        })
        .catch((err) => XLogger.error("Failed to search drawings", err));
    }, 250);

    return () => clearTimeout(timer);
  }, [search, isOpen]);

  const handleSave = useCallback(async () => {
    if (!currentDrawingId || saving) return;

    setSaving(true);
    try {
      const drawingDataState = await getDrawingDataState();
      await browser.runtime.sendMessage(
        As<SaveDrawingMessage>({
          type: MessageType.SAVE_DRAWING,
          payload: {
            id: currentDrawingId,
            excalidraw: drawingDataState.excalidraw,
            excalidrawState: drawingDataState.excalidrawState,
            versionFiles: drawingDataState.versionFiles,
            versionDataState: drawingDataState.versionDataState,
            imageBase64: drawingDataState.imageBase64,
            viewBackgroundColor: drawingDataState.viewBackgroundColor,
          },
        })
      );
    } catch (err) {
      XLogger.error("Failed to save drawing", err);
    } finally {
      setSaving(false);
    }
  }, [currentDrawingId, saving]);

  const handleSaveNew = useCallback(async () => {
    if (saving || !newName.trim()) return;

    setSaving(true);
    try {
      await browser.runtime.sendMessage({
        type: MessageType.MESSAGE_AUTO_SAVE,
        payload: { name: newName.trim(), setCurrent: true },
      });
      setIsOpen(false);
    } catch (err) {
      XLogger.error("Failed to save new drawing", err);
    } finally {
      setSaving(false);
    }
  }, [saving, newName]);

  const handleLoadDrawing = useCallback(async (id: string) => {
    setIsOpen(false);
    try {
      await browser.runtime.sendMessage({
        type: MessageType.LOAD_DRAWING,
        payload: { id },
      });
    } catch (err) {
      XLogger.error("Failed to load drawing", err);
    }
  }, []);

  const handleNewDrawing = useCallback(async () => {
    setIsOpen(false);
    try {
      await browser.runtime.sendMessage({
        type: MessageType.CREATE_NEW_DRAWING,
      });
    } catch (err) {
      XLogger.error("Failed to create new drawing", err);
    }
  }, []);

  const handleOpenFullUI = useCallback(async () => {
    setIsOpen(false);
    try {
      await browser.runtime.sendMessage({
        type: MessageType.OPEN_POPUP,
      });
    } catch (err) {
      XLogger.error("Failed to open full UI", err);
    }
  }, []);

  return (
    <>
      <div className="excalisave-title-wrapper">
        <h1
          className="excalisave-title"
          onMouseEnter={() => setShowTooltip(true)}
          onMouseLeave={() => setShowTooltip(false)}
        >
          {title}
        </h1>
        {showTooltip && title && (
          <div className="excalisave-title-tooltip">{title}</div>
        )}
      </div>
      <div ref={dropdownRef} style={{ position: "relative", marginLeft: "8px" }}>
        <button
          className="excalidraw-button collab-button excalisave-button"
          style={{ width: "auto" }}
          title="Open Excalisave"
          onClick={() => setIsOpen(!isOpen)}
        >
          Excalisave
        </button>
        {isOpen && (
          <div className="excalisave-dropdown">
            <div className="excalisave-dropdown__actions">
              {isUnsaved ? (
                <div className="excalisave-dropdown__save-new">
                  <input
                    className="excalisave-dropdown__name-input"
                    type="text"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleSaveNew()}
                    placeholder="Drawing name"
                    autoFocus
                  />
                  <button
                    className="excalisave-dropdown__action-btn"
                    onClick={handleSaveNew}
                    disabled={saving || !newName.trim()}
                  >
                    {saving ? "Saving..." : "Save"}
                  </button>
                </div>
              ) : (
                <>
                  <button
                    className="excalisave-dropdown__action-btn"
                    onClick={handleSave}
                    disabled={saving}
                  >
                    {saving ? "Saving..." : "Save"}
                  </button>
                  <button
                    className="excalisave-dropdown__action-btn"
                    onClick={handleNewDrawing}
                  >
                    New Drawing
                  </button>
                </>
              )}
            </div>
            <div className="excalisave-dropdown__divider" />
            {!loading && drawings.length > 0 && (
              <div className="excalisave-dropdown__search">
                <input
                  className="excalisave-dropdown__name-input"
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search drawings..."
                />
              </div>
            )}
            <div className="excalisave-dropdown__list">
              {loading ? (
                <div className="excalisave-dropdown__empty">Loading...</div>
              ) : drawings.length === 0 ? (
                <div className="excalisave-dropdown__empty">
                  No drawings saved
                </div>
              ) : (
                (() => {
                  const displayList = searchResults !== null ? searchResults : drawings;
                  return displayList.length === 0 ? (
                    <div className="excalisave-dropdown__empty">
                      No matches
                    </div>
                  ) : (
                    displayList.map((drawing) => (
                      <button
                        key={drawing.id}
                        className={`excalisave-dropdown__item ${
                          drawing.id === currentDrawingId
                            ? "excalisave-dropdown__item--active"
                            : ""
                        }`}
                        onClick={() => handleLoadDrawing(drawing.id)}
                        title={drawing.name}
                      >
                        {drawing.roomUrl && (
                          <span className="excalisave-dropdown__shared-icon" title="Shared session">
                            <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                              <circle cx="8" cy="8" r="7" fill="#22c55e" />
                              <path d="M5.5 8.5L7 10L10.5 6.5" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                            </svg>
                          </span>
                        )}
                        <span className="excalisave-dropdown__item-name">{drawing.name}</span>
                      </button>
                    ))
                  );
                })()
              )}
            </div>
            <div className="excalisave-dropdown__divider" />
            <div className="excalisave-dropdown__footer">
              <button
                className="excalisave-dropdown__footer-btn"
                onClick={handleOpenFullUI}
              >
                Open Full UI
              </button>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
