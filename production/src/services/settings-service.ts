declare const Office: any;

export interface IssueFlowSettings {
  parties: string[];
  statuses: string[];
}

const SETTINGS_KEY = "issueflow.settings.v1";

const DEFAULT_SETTINGS: IssueFlowSettings = {
  parties: ["Architect", "C&S", "MEP", "Main Contractor", "ESCS"],
  statuses: ["Open", "In Progress", "Pending", "Closed"],
};

function normalizeList(values: unknown, fallback: string[]): string[] {
  if (!Array.isArray(values)) return [...fallback];

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

  return result.length ? result : [...fallback];
}

export function loadSettings(): IssueFlowSettings {
  const stored = Office.context.document.settings.get(SETTINGS_KEY);

  if (!stored || typeof stored !== "object") {
    return {
      parties: [...DEFAULT_SETTINGS.parties],
      statuses: [...DEFAULT_SETTINGS.statuses],
    };
  }

  return {
    parties: normalizeList(stored.parties, DEFAULT_SETTINGS.parties),
    statuses: normalizeList(stored.statuses, DEFAULT_SETTINGS.statuses),
  };
}

export function saveSettings(settings: IssueFlowSettings): Promise<void> {
  const normalized: IssueFlowSettings = {
    parties: normalizeList(settings.parties, DEFAULT_SETTINGS.parties),
    statuses: normalizeList(settings.statuses, DEFAULT_SETTINGS.statuses),
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
