export enum SyncStatus {
  SYNCED = "SYNCED",
  SYNCING = "SYNCING",
  PENDING = "PENDING",
  ERROR = "ERROR",
  CONFLICT = "CONFLICT",
}

export interface SyncStatusInfo {
  status: SyncStatus;
  lastSyncedAt?: Date;
  error?: string;
  progress?: number;
  version: string;
}

export interface FileDelta {
  path: string;
  changes: {
    added: string[];
    modified: string[];
    deleted: string[];
  };
  baseVersion: string;
  targetVersion: string;
}

export interface SyncStatusIndicatorProps {
  status: SyncStatus;
  lastSyncedAt?: Date;
  progress?: number;
  error?: string;
}
