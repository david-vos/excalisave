import React from "react";
import { Dialog, Box, Text, Button, Theme } from "@radix-ui/themes";
import { IDrawing } from "../../interfaces/drawing.interface";
import "./SyncConflictDialog.styles.scss";

interface SyncConflictDialogProps {
  isOpen: boolean;
  onClose: () => void;
  localDrawing: IDrawing;
  oneDriveDrawing: IDrawing;
  onResolve: (useLocal: boolean) => void;
}

/**
 * Dialog component for resolving sync conflicts between local and cloud versions
 */
export const SyncConflictDialog: React.FC<SyncConflictDialogProps> = ({
  isOpen,
  onClose,
  localDrawing,
  oneDriveDrawing,
  onResolve,
}) => {
  if (!isOpen) return null;

  const formatDate = (date: string) => {
    return new Date(date).toLocaleString();
  };

  return (
    <Theme accentColor="iris">
      <Dialog.Root open={isOpen} onOpenChange={onClose}>
        <Dialog.Content>
          <Dialog.Title>Sync Conflict Detected</Dialog.Title>
          <Text as="p" size="2" mb="4">
            There are conflicting changes between your local version and the
            OneDrive version of "{localDrawing.name}". Which version would you
            like to keep?
          </Text>

          <Box className="sync-conflict-dialog__content">
            <Box className="sync-conflict-dialog__version">
              <Text as="div" size="2" weight="bold">
                Local Version
              </Text>
              <Text as="div" size="1" color="gray">
                Last modified: {formatDate(localDrawing.createdAt)}
              </Text>
            </Box>

            <Box className="sync-conflict-dialog__version">
              <Text as="div" size="2" weight="bold">
                OneDrive Version
              </Text>
              <Text as="div" size="1" color="gray">
                Last modified: {formatDate(oneDriveDrawing.createdAt)}
              </Text>
            </Box>
          </Box>

          <Box className="sync-conflict-dialog__actions">
            <Button variant="soft" color="red" onClick={() => onResolve(true)}>
              Use Local Version
            </Button>
            <Button
              variant="soft"
              color="blue"
              onClick={() => onResolve(false)}
            >
              Use OneDrive Version
            </Button>
          </Box>
        </Dialog.Content>
      </Dialog.Root>
    </Theme>
  );
};
