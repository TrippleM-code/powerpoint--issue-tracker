import type { Issue } from "./models";

export interface ValidationResult {
  valid: boolean;
  errors: string[];
}

export function validateIssue(issue: Issue): ValidationResult {
  const errors: string[] = [];

  if (!issue.id.trim()) errors.push("Issue ID is required.");
  if (!issue.areaCode.trim()) errors.push("Area Code is required.");
  if (!issue.roomSpace.trim()) errors.push("Room / Space is required.");
  if (!issue.description.trim()) errors.push("Description is required.");

  issue.actions.forEach((action, index) => {
    if (!action.party.trim()) errors.push(`Action ${index + 1}: party is required.`);
    if (!action.required.trim()) errors.push(`Action ${index + 1}: action text is required.`);
    if (!action.status.trim()) errors.push(`Action ${index + 1}: status is required.`);
  });

  return { valid: errors.length === 0, errors };
}
