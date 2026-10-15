export interface TemplateSettings {
  format: "standard" | "ice";
  title: string;
  titles: string[];
  meetingNumber: string;
  meetingDate: string;
}
export const TEMPLATE_KEY = "issueflow.template.v1";
const text = (value: unknown, max: number) => typeof value === "string" ? value.trim().slice(0, max) : "";
export function normalizeTemplate(value: unknown): TemplateSettings {
  const v = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const date = text(v.meetingDate, 10);
  return {
    format: v.format === "ice" ? "ice" : "standard",
    title: text(v.title, 100),
    titles: Array.isArray(v.titles) ? [...new Set(v.titles.map(t => text(t, 100)).filter(Boolean))].slice(-30) : [],
    meetingNumber: text(v.meetingNumber, 12),
    meetingDate: /^\d{4}-\d{2}-\d{2}$/.test(date) && !Number.isNaN(Date.parse(date)) ? date : "",
  };
}
export function loadTemplate(): TemplateSettings {
  return normalizeTemplate(typeof Office === "undefined" ? undefined : Office.context?.document?.settings?.get(TEMPLATE_KEY));
}
export async function saveTemplate(value: TemplateSettings): Promise<void> {
  const normalized = normalizeTemplate(value);
  if (normalized.format === "ice" && !normalized.title) throw new Error("Enter or choose a project / meeting title.");
  if (normalized.title) normalized.titles = [...new Set([...normalized.titles, normalized.title])].slice(-30);
  const store = Office.context.document.settings;
  const previous = store.get(TEMPLATE_KEY);
  store.set(TEMPLATE_KEY, normalized);
  await new Promise<void>((resolve, reject) => store.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else {
      if (previous === undefined) store.remove(TEMPLATE_KEY); else store.set(TEMPLATE_KEY, previous);
      reject(new Error(result.error?.message || "Could not save template settings."));
    }
  }));
}
