export const DEFAULT_STATUSES = [
  "Open",
  "In Progress",
  "Pending",
  "Closed",
] as const;

export function computeOverallStatus(statuses: string[]): string {
  if (statuses.length === 0) return "No actions";
  const normalized = statuses.map((s) => s.trim().toLowerCase());

  if (normalized.every((s) => s === "closed")) return "Closed";
  if (normalized.some((s) => s === "open")) return "Open";
  if (normalized.some((s) => s === "in progress")) return "In Progress";
  if (normalized.some((s) => s === "pending")) return "Pending";
  return "Open";
}
