import { TAGS } from "../storage/tag-names";
import type { Issue } from "../domain/models";

export const REFERENCE_ROLES = ["SOI", "NOV", "COI/CVI"] as const;
export type ReferenceRole = typeof REFERENCE_ROLES[number];
export type IssueReferences = Partial<Record<ReferenceRole, string>>;

/** Read live user-owned reference text; never copy it into issue metadata or notes. */
export async function readIssueReferences(
  context: any, records: {slideId: string; issue: Issue}[],
): Promise<Map<Issue, IssueReferences>> {
  const entries = records.map(record => ({...record, slide: context.presentation.slides.getItem(record.slideId)}));
  for (const entry of entries) entry.slide.shapes.load("items/id");
  await context.sync();
  for (const entry of entries) for (const shape of entry.slide.shapes.items) shape.tags.load("items/key,value");
  await context.sync();
  const fields: {issue: Issue; role: ReferenceRole; range: any}[] = [];
  for (const entry of entries) for (const shape of entry.slide.shapes.items) {
    const role = shape.tags.items.find((tag: any) => tag.key === TAGS.referenceRole)?.value;
    if (!REFERENCE_ROLES.includes(role)) continue;
    const range = shape.textFrame.textRange;
    range.load("text");
    fields.push({issue: entry.issue, role, range});
  }
  await context.sync();
  const result = new Map<Issue, IssueReferences>();
  for (const field of fields) {
    const value = field.range.text.trim();
    if (!value) continue;
    const refs = result.get(field.issue) ?? {};
    refs[field.role] = [...new Set([...(refs[field.role]?.split("\n") ?? []), value])].join("\n");
    result.set(field.issue, refs);
  }
  return result;
}
