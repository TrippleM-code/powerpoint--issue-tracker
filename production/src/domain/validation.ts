import type { Issue } from "./models";

export type IssueDraft = Pick<Issue, "id" | "areaCode" | "roomSpace" | "description">;

export function normalizeIssueId(value: string): string {
  return value.trim().toUpperCase();
}

export function validateIssueDraft(draft: IssueDraft): string[] {
  const errors: string[] = [];
  if (!draft.id.trim()) errors.push("Issue ID is required.");
  if (!draft.areaCode.trim()) errors.push("Area Code is required.");
  if (!draft.roomSpace.trim()) errors.push("Room / Space is required.");
  if (!draft.description.trim()) errors.push("Description is required.");
  return errors;
}

export function isDuplicateIssueId(candidateId: string, existingIds: string[]): boolean {
  const candidate = normalizeIssueId(candidateId);
  if (!candidate) return false;

  return existingIds.some((existingId) => normalizeIssueId(existingId) === candidate);
}
