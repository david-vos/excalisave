import React, { useMemo } from "react";
import { Box, Text, Theme } from "@radix-ui/themes";
import {
  SyncStatus,
  SyncStatusIndicatorProps,
} from "../../interfaces/sync-status.interface";
import {
  getStatusConfig,
  formatLastSynced,
} from "../../utils/sync-status.utils";
import "./SyncStatusIndicator.styles.scss";

export const SyncStatusIndicator: React.FC<SyncStatusIndicatorProps> = ({
  status,
  lastSyncedAt,
  progress,
  error,
}) => {
  const statusConfig = useMemo(() => getStatusConfig(status), [status]);
  const formattedLastSynced = useMemo(
    () => formatLastSynced(lastSyncedAt),
    [lastSyncedAt]
  );

  const progressValue = Math.min(Math.max(progress || 0, 0), 100);

  return (
    <Theme accentColor="iris">
      <Box
        className="sync-status-indicator"
        data-status={status.toLowerCase()}
        role="status"
        aria-label={`Sync status: ${statusConfig.label}`}
      >
        <Box
          className="sync-status-indicator__icon"
          style={{ color: `var(--${statusConfig.color}-9)` }}
          aria-hidden="true"
        >
          {statusConfig.icon}
        </Box>
        <Box className="sync-status-indicator__details">
          <Text size="2" weight="medium">
            {statusConfig.label}
          </Text>
          {lastSyncedAt && (
            <Text size="1" color="gray">
              Last synced: {formattedLastSynced}
            </Text>
          )}
          {error && (
            <Text size="1" color="red" role="alert">
              {error}
            </Text>
          )}
          {progress !== undefined && status === SyncStatus.SYNCING && (
            <Box
              className="sync-status-indicator__progress"
              role="progressbar"
              aria-valuenow={progressValue}
              aria-valuemin={0}
              aria-valuemax={100}
            >
              <Box
                className="sync-status-indicator__progress-bar"
                style={{ width: `${progressValue}%` }}
              />
            </Box>
          )}
        </Box>
      </Box>
    </Theme>
  );
};
