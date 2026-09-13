import type { Issue } from "../domain/models";
import { normalizeIssueId, validateIssueDraft } from "../domain/validation";
import { PowerPointService } from "../services/powerpoint-service";

const service = new PowerPointService();

type Elements = {
  issueId: HTMLInputElement;
  areaCode: HTMLInputElement;
  roomSpace: HTMLInputElement;
  description: HTMLTextAreaElement;
  save: HTMLButtonElement;
  refresh: HTMLButtonElement;
  banner: HTMLElement;
  createdAt: HTMLElement;
  updatedAt: HTMLElement;
};

let currentIssue: Issue | null = null;

function el<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`Missing UI element: ${id}`);
  return found as T;
}

function elements(): Elements {
  return {
    issueId: el<HTMLInputElement>("issueId"),
    areaCode: el<HTMLInputElement>("areaCode"),
    roomSpace: el<HTMLInputElement>("roomSpace"),
    description: el<HTMLTextAreaElement>("description"),
    save: el<HTMLButtonElement>("saveIssueBtn"),
    refresh: el<HTMLButtonElement>("refreshSheetBtn"),
    banner: el<HTMLElement>("statusBanner"),
    createdAt: el<HTMLElement>("createdAt"),
    updatedAt: el<HTMLElement>("updatedAt"),
  };
}

function showBanner(message: string, type: "success" | "error" | "info"): void {
  const ui = elements();
  ui.banner.textContent = message;
  ui.banner.className = `banner ${type}`;
}

function setBusy(busy: boolean): void {
  const ui = elements();
  ui.save.disabled = busy;
  ui.refresh.disabled = busy;
}

function formatTimestamp(value?: string): string {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function populate(issue: Issue | null): void {
  const ui = elements();
  currentIssue = issue;

  ui.issueId.value = issue?.id ?? "";
  ui.areaCode.value = issue?.areaCode ?? "";
  ui.roomSpace.value = issue?.roomSpace ?? "";
  ui.description.value = issue?.description ?? "";
  ui.createdAt.textContent = formatTimestamp(issue?.createdAt);
  ui.updatedAt.textContent = formatTimestamp(issue?.updatedAt);
}

function buildIssue(): Issue {
  const ui = elements();
  const now = new Date().toISOString();

  const draft = {
    id: normalizeIssueId(ui.issueId.value),
    areaCode: ui.areaCode.value.trim(),
    roomSpace: ui.roomSpace.value.trim(),
    description: ui.description.value.trim(),
  };

  const errors = validateIssueDraft(draft);
  if (errors.length) throw new Error(errors.join(" "));

  const contentChanged =
    !currentIssue ||
    currentIssue.id !== draft.id ||
    currentIssue.areaCode !== draft.areaCode ||
    currentIssue.roomSpace !== draft.roomSpace ||
    currentIssue.description !== draft.description;

  return {
    ...draft,
    createdAt: currentIssue?.createdAt ?? now,
    updatedAt: currentIssue && contentChanged ? now : currentIssue?.updatedAt,
    actions: currentIssue?.actions ?? [],
  };
}

async function saveOnly(): Promise<void> {
  setBusy(true);
  try {
    const issue = buildIssue();
    await service.saveSelectedIssue(issue);
    populate(issue);
    showBanner("Issue data saved.", "success");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function refreshSheet(): Promise<void> {
  setBusy(true);
  try {
    const issue = buildIssue();
    await service.saveSelectedIssue(issue);
    await service.renderSelectedIssue(issue);
    populate(issue);
    showBanner("Issue sheet refreshed. Manual reference content was preserved.", "success");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  } finally {
    setBusy(false);
  }
}

export async function initializeIssuePanel(): Promise<void> {
  elements().save.addEventListener("click", saveOnly);
  elements().refresh.addEventListener("click", refreshSheet);

  try {
    const issue = await service.readSelectedIssue();
    populate(issue);
    if (issue) showBanner("Existing IssueFlow issue loaded.", "info");
  } catch (error) {
    populate(null);
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}
