export const DEFAULT_STATUSES = [
  "Open",
  "In Progress",
  "Pending",
  "Closed",
] as const;

export function computeOverallStatus(statuses: string[]): string {
  if (statuses.length === 0) return "No actions";

  const normalized = statuses.map((status) => status.trim().toLowerCase());

  if (normalized.every((status) => status === "closed")) return "Closed";
  if (normalized.some((status) => status === "open")) return "Open";
  if (normalized.some((status) => status === "in progress")) return "In Progress";
  if (normalized.some((status) => status === "pending")) return "Pending";

  // Unknown/custom statuses remain conservative.
  return "Open";
}

export function statusCssClass(status: string): string {
  switch (status.trim().toLowerCase()) {
    case "open": return "status-open";
    case "in progress": return "status-in-progress";
    case "pending": return "status-pending";
    case "closed": return "status-closed";
    default: return "status-custom";
  }
}
