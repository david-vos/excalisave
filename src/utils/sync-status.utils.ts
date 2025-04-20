import { SyncStatus } from "../interfaces/sync-status.interface";

export const getStatusConfig = (status: SyncStatus) => {
  const configs = {
    [SyncStatus.SYNCED]: { color: "green", icon: "✓", label: "Synced" },
    [SyncStatus.SYNCING]: { color: "blue", icon: "↻", label: "Syncing" },
    [SyncStatus.PENDING]: { color: "yellow", icon: "⋯", label: "Pending" },
    [SyncStatus.ERROR]: { color: "red", icon: "!", label: "Error" },
    [SyncStatus.CONFLICT]: { color: "orange", icon: "⚠", label: "Conflict" },
  };
  return configs[status] || { color: "gray", icon: "?", label: "Unknown" };
};

export const formatLastSynced = (date?: Date): string => {
  if (!date) return "";
  const now = new Date();
  const diff = now.getTime() - date.getTime();
  const minutes = Math.floor(diff / 60000);

  if (minutes < 1) return "Just now";
  if (minutes < 60) return `${minutes}m ago`;
  if (minutes < 1440) return `${Math.floor(minutes / 60)}h ago`;
  return date.toLocaleDateString();
};
