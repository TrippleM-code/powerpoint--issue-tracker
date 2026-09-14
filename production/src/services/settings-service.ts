import type { Party } from "../domain/models";

declare const Office: any;

export interface IssueFlowSettings {
  parties: Party[];
  statuses: string[];
}

const SETTINGS_KEY = "issueflow.settings.v1";
const DEFAULT_PARTY_NAMES = ["Architect", "C&S", "MEP", "Main Contractor", "ESCS"];
const DEFAULT_STATUSES = ["Open", "In Progress", "Pending", "Closed"];

function slugId(name: string, index: number): string {
  const slug = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return `party-${slug || "item"}-${index + 1}`;
}

function normalizeStatuses(values: unknown): string[] {
  if (!Array.isArray(values)) return [...DEFAULT_STATUSES];
  const result: string[] = [];
  const seen = new Set<string>();
  for (const value of values) {
    const text = String(value ?? "").trim();
    if (!text) continue;
    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(text);
  }
  return result.length ? result : [...DEFAULT_STATUSES];
}

function normalizeParties(values: unknown): Party[] {
  if (!Array.isArray(values)) {
    return DEFAULT_PARTY_NAMES.map((name, index) => ({ id: slugId(name, index), name }));
  }

  const result: Party[] = [];
  const seen = new Set<string>();

  values.forEach((value, index) => {
    let name = "";
    let id = "";
    let logoDataUrl: string | undefined;

    if (typeof value === "string") {
      name = value.trim();
      id = slugId(name, index);
    } else if (value && typeof value === "object") {
      const candidate = value as Record<string, unknown>;
      name = String(candidate.name ?? "").trim();
      id = String(candidate.id ?? "").trim() || slugId(name, index);
      const logo = String(candidate.logoDataUrl ?? "").trim();
      logoDataUrl = logo || undefined;
    }

    if (!name) return;
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);

    result.push({ id, name, ...(logoDataUrl ? { logoDataUrl } : {}) });
  });

  return result.length
    ? result
    : DEFAULT_PARTY_NAMES.map((name, index) => ({ id: slugId(name, index), name }));
}

export function loadSettings(): IssueFlowSettings {
  const stored = Office.context.document.settings.get(SETTINGS_KEY);
  if (!stored || typeof stored !== "object") {
    return { parties: normalizeParties(undefined), statuses: [...DEFAULT_STATUSES] };
  }
  return {
    parties: normalizeParties(stored.parties),
    statuses: normalizeStatuses(stored.statuses),
  };
}

export function saveSettings(settings: IssueFlowSettings): Promise<void> {
  const normalized: IssueFlowSettings = {
    parties: normalizeParties(settings.parties),
    statuses: normalizeStatuses(settings.statuses),
  };

  Office.context.document.settings.set(SETTINGS_KEY, normalized);

  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message || "Could not save IssueFlow settings."));
      }
    });
  });
}
