import type { Issue } from "./models";

export const CONFLICT_MESSAGE = "This issue changed since it was loaded. Your draft has not been saved. Copy any edits you want to keep, then select another slide and return to reload the latest issue.";

// Compare content, not just timestamps; imported/older editors may not update them.
export function issueRevision(issue: Issue | null): string {
  if (!issue) return "null";
  return JSON.stringify([issue.id, issue.areaCode, issue.roomSpace, issue.description,
    issue.createdAt, issue.updatedAt ?? null,
    issue.actions.map(a => [a.id, a.party, a.required, a.status, a.createdAt, a.updatedAt ?? null])]);
}
