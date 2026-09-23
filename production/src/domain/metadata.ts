import type { Issue } from "./models";
import { SCHEMA_VERSION } from "../storage/tag-names";

function object(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function text(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function date(value: unknown): boolean {
  return text(value) && Number.isFinite(Date.parse(value));
}

export function validateIssue(value: unknown): asserts value is Issue {
  if (!object(value) || ![value.id, value.areaCode, value.roomSpace, value.description].every(text)
      || !date(value.createdAt) || (value.updatedAt !== undefined && !date(value.updatedAt))
      || !Array.isArray(value.actions)) {
    throw new Error("Invalid issue fields, timestamps, or actions list.");
  }
  const ids = new Set<string>();
  for (const action of value.actions) {
    if (!object(action) || ![action.id, action.party, action.required, action.status].every(text)
        || !date(action.createdAt) || (action.updatedAt !== undefined && !date(action.updatedAt))) {
      throw new Error("Invalid action fields or timestamps.");
    }
    const id = action.id as string;
    if (ids.has(id)) throw new Error(`Duplicate action ID: ${id}.`);
    ids.add(id);
  }
}

export function parseIssueMetadata(json: string, schema: string | undefined, location: string): Issue {
  try {
    // Legacy v1 slides may not have a schema tag. Never reinterpret future schemas.
    if (schema && schema !== SCHEMA_VERSION) throw new Error(`Unsupported schema version ${schema}.`);
    const value: unknown = JSON.parse(json);
    validateIssue(value);
    return value;
  } catch (error) {
    throw new Error(`${location}: invalid IssueFlow metadata. ${error instanceof Error ? error.message : String(error)}`);
  }
}
