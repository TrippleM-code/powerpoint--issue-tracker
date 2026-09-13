import type { ActionItem, Issue } from "../domain/models";
import { normalizeIssueId, validateIssueDraft } from "../domain/validation";
import { statusCssClass } from "../domain/status";
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
  editingActionId: HTMLInputElement;
  actionParty: HTMLInputElement;
  actionRequired: HTMLTextAreaElement;
  actionStatus: HTMLSelectElement;
  saveAction: HTMLButtonElement;
  cancelEditAction: HTMLButtonElement;
  actionList: HTMLElement;
};

let currentIssue: Issue | null = null;
let removeArmedId: string | null = null;

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
    editingActionId: el<HTMLInputElement>("editingActionId"),
    actionParty: el<HTMLInputElement>("actionParty"),
    actionRequired: el<HTMLTextAreaElement>("actionRequired"),
    actionStatus: el<HTMLSelectElement>("actionStatus"),
    saveAction: el<HTMLButtonElement>("saveActionBtn"),
    cancelEditAction: el<HTMLButtonElement>("cancelEditActionBtn"),
    actionList: el<HTMLElement>("actionList"),
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
  ui.saveAction.disabled = busy;
}

function formatTimestamp(value?: string): string {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function makeActionId(): string {
  const suffix = Math.random().toString(36).slice(2, 8).toUpperCase();
  return `ACT-${Date.now().toString(36).toUpperCase()}-${suffix}`;
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

  renderActionList();
}

function buildIssueFromForm(): Issue {
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

async function persistIssue(issue: Issue, message: string): Promise<void> {
  await service.saveSelectedIssue(issue);
  populate(issue);
  showBanner(message, "success");
}

async function saveOnly(): Promise<void> {
  setBusy(true);
  try {
    const issue = buildIssueFromForm();
    await persistIssue(issue, "Issue data saved.");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function refreshSheet(): Promise<void> {
  setBusy(true);
  try {
    const issue = buildIssueFromForm();
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

function resetActionEditor(): void {
  const ui = elements();
  ui.editingActionId.value = "";
  ui.actionParty.value = "";
  ui.actionRequired.value = "";
  ui.actionStatus.value = "Open";
  ui.saveAction.textContent = "Add Action";
  ui.cancelEditAction.classList.add("hidden");
  removeArmedId = null;
}

function editAction(actionId: string): void {
  if (!currentIssue) return;
  const action = currentIssue.actions.find((item) => item.id === actionId);
  if (!action) return;

  const ui = elements();
  ui.editingActionId.value = action.id;
  ui.actionParty.value = action.party;
  ui.actionRequired.value = action.required;
  ui.actionStatus.value = action.status;
  ui.saveAction.textContent = "Save Action";
  ui.cancelEditAction.classList.remove("hidden");
  ui.actionParty.focus();
}

async function saveAction(): Promise<void> {
  setBusy(true);
  try {
    const ui = elements();
    const baseIssue = buildIssueFromForm();
    const party = ui.actionParty.value.trim();
    const required = ui.actionRequired.value.trim();
    const status = ui.actionStatus.value.trim();

    if (!party) throw new Error("Action By / Responsible Party is required.");
    if (!required) throw new Error("Action Required is required.");
    if (!status) throw new Error("Status is required.");

    const now = new Date().toISOString();
    const editingId = ui.editingActionId.value.trim();

    let actions: ActionItem[];
    if (editingId) {
      actions = baseIssue.actions.map((action) =>
        action.id === editingId
          ? { ...action, party, required, status, updatedAt: now }
          : action
      );
    } else {
      actions = [
        ...baseIssue.actions,
        {
          id: makeActionId(),
          party,
          required,
          status,
          createdAt: now,
        },
      ];
    }

    const issue: Issue = {
      ...baseIssue,
      actions,
      updatedAt: now,
    };

    await service.saveSelectedIssue(issue);
    await service.renderSelectedIssue(issue);
    populate(issue);
    resetActionEditor();
    showBanner(editingId ? "Action updated." : "Action added.", "success");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  } finally {
    setBusy(false);
  }
}

async function removeAction(actionId: string): Promise<void> {
  if (!currentIssue) return;

  if (removeArmedId !== actionId) {
    removeArmedId = actionId;
    renderActionList();
    showBanner("Click Remove again to confirm.", "info");
    return;
  }

  setBusy(true);
  try {
    const now = new Date().toISOString();
    const issue: Issue = {
      ...currentIssue,
      actions: currentIssue.actions.filter((action) => action.id !== actionId),
      updatedAt: now,
    };

    await service.saveSelectedIssue(issue);
    await service.renderSelectedIssue(issue);
    populate(issue);
    resetActionEditor();
    showBanner("Action removed.", "success");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  } finally {
    setBusy(false);
  }
}

function renderActionList(): void {
  const ui = elements();
  ui.actionList.innerHTML = "";

  const actions = currentIssue?.actions ?? [];
  if (actions.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No actions yet.";
    ui.actionList.appendChild(empty);
    return;
  }

  actions.forEach((action) => {
    const card = document.createElement("article");
    card.className = "action-card";

    const head = document.createElement("div");
    head.className = "action-card-head";

    const title = document.createElement("div");
    title.className = "action-card-title";
    title.textContent = action.party;

    const status = document.createElement("span");
    status.className = `action-card-status ${statusCssClass(action.status)}`;
    status.textContent = action.status;

    head.append(title, status);

    const text = document.createElement("p");
    text.textContent = action.required;

    const buttons = document.createElement("div");
    buttons.className = "action-card-actions";

    const edit = document.createElement("button");
    edit.className = "secondary";
    edit.textContent = "Edit";
    edit.addEventListener("click", () => editAction(action.id));

    const remove = document.createElement("button");
    remove.className = "secondary";
    remove.textContent = removeArmedId === action.id ? "Confirm Remove" : "Remove";
    remove.addEventListener("click", () => removeAction(action.id));

    buttons.append(edit, remove);
    card.append(head, text, buttons);
    ui.actionList.appendChild(card);
  });
}

export async function initializeIssuePanel(): Promise<void> {
  const ui = elements();
  ui.save.addEventListener("click", saveOnly);
  ui.refresh.addEventListener("click", refreshSheet);
  ui.saveAction.addEventListener("click", saveAction);
  ui.cancelEditAction.addEventListener("click", resetActionEditor);

  try {
    const issue = await service.readSelectedIssue();
    populate(issue);
    if (issue) showBanner("Existing IssueFlow issue loaded.", "info");
  } catch (error) {
    populate(null);
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}
