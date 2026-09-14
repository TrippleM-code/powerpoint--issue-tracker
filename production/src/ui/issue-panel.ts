import type { ActionItem, Issue, Party } from "../domain/models";
import { normalizeIssueId, validateIssueDraft } from "../domain/validation";
import { statusCssClass } from "../domain/status";
import { PowerPointService } from "../services/powerpoint-service";
import {
  loadSettings,
  saveSettings,
  type IssueFlowSettings,
} from "../services/settings-service";

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
  actionParty: HTMLSelectElement;
  actionRequired: HTMLTextAreaElement;
  actionStatus: HTMLSelectElement;
  saveAction: HTMLButtonElement;
  cancelEditAction: HTMLButtonElement;
  actionList: HTMLElement;
  newPartyName: HTMLInputElement;
  addParty: HTMLButtonElement;
  partyLibraryList: HTMLElement;
  newStatusName: HTMLInputElement;
  addStatus: HTMLButtonElement;
  statusLibraryList: HTMLElement;
};

let currentIssue: Issue | null = null;
let settings: IssueFlowSettings = loadSettings();
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
    actionParty: el<HTMLSelectElement>("actionParty"),
    actionRequired: el<HTMLTextAreaElement>("actionRequired"),
    actionStatus: el<HTMLSelectElement>("actionStatus"),
    saveAction: el<HTMLButtonElement>("saveActionBtn"),
    cancelEditAction: el<HTMLButtonElement>("cancelEditActionBtn"),
    actionList: el<HTMLElement>("actionList"),
    newPartyName: el<HTMLInputElement>("newPartyName"),
    addParty: el<HTMLButtonElement>("addPartyBtn"),
    partyLibraryList: el<HTMLElement>("partyLibraryList"),
    newStatusName: el<HTMLInputElement>("newStatusName"),
    addStatus: el<HTMLButtonElement>("addStatusBtn"),
    statusLibraryList: el<HTMLElement>("statusLibraryList"),
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
  ui.addParty.disabled = busy;
  ui.addStatus.disabled = busy;
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

function fillSelect(select: HTMLSelectElement, values: string[], preferred?: string): void {
  select.innerHTML = "";

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  if (preferred && values.some((value) => value.toLowerCase() === preferred.toLowerCase())) {
    select.value = values.find((value) => value.toLowerCase() === preferred.toLowerCase()) || values[0] || "";
  } else {
    select.value = values[0] || "";
  }
}

function refreshActionChoices(preferredParty?: string, preferredStatus?: string): void {
  const ui = elements();
  fillSelect(ui.actionParty, settings.parties.map((party) => party.name), preferredParty);
  fillSelect(ui.actionStatus, settings.statuses, preferredStatus);
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
  ui.actionRequired.value = "";
  refreshActionChoices();
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

  // If an old action uses a value no longer in the library, do not silently
  // replace it. Ask the user to restore that value in Settings first.
  const partyExists = settings.parties.some(
    (value) => value.toLowerCase() === action.party.toLowerCase()
  );
  const statusExists = settings.statuses.some(
    (value) => value.toLowerCase() === action.status.toLowerCase()
  );

  if (!partyExists || !statusExists) {
    showBanner(
      "This action uses a party or status that is no longer in Settings. Add it back before editing.",
      "error"
    );
    return;
  }

  refreshActionChoices(action.party, action.status);
  ui.actionRequired.value = action.required;
  ui.saveAction.textContent = "Save Action";
  ui.cancelEditAction.classList.remove("hidden");
  ui.actionRequired.focus();
}

async function saveAction(): Promise<void> {
  setBusy(true);
  try {
    const ui = elements();
    const baseIssue = buildIssueFromForm();
    const party = ui.actionParty.value.trim();
    const required = ui.actionRequired.value.trim();
    const status = ui.actionStatus.value.trim();

    if (!party) throw new Error("Choose an Action By / Responsible Party.");
    if (!required) throw new Error("Action Required is required.");
    if (!status) throw new Error("Choose a Status.");

    if (!settings.parties.includes(party)) {
      throw new Error("Selected party is not in the Settings party library.");
    }
    if (!settings.statuses.includes(status)) {
      throw new Error("Selected status is not in the Settings status library.");
    }

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

function initializeTabs(): void {
  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>(".tab-button"));
  const panels = Array.from(document.querySelectorAll<HTMLElement>(".tab-panel"));

  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.tab;

      buttons.forEach((item) => {
        const active = item === button;
        item.classList.toggle("active", active);
        item.setAttribute("aria-selected", String(active));
      });

      panels.forEach((panel) => {
        panel.classList.toggle("active", panel.id === `tab-${target}`);
      });
    });
  });
}

function renderLibraryList(
  container: HTMLElement,
  values: string[],
  onRemove: (value: string) => void
): void {
  container.innerHTML = "";

  values.forEach((value) => {
    const item = document.createElement("div");
    item.className = "library-item";

    const name = document.createElement("span");
    name.className = "library-name";
    name.textContent = value;

    const remove = document.createElement("button");
    remove.className = "library-remove";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => onRemove(value));

    item.append(name, remove);
    container.appendChild(item);
  });
}

async function persistSettings(message: string): Promise<void> {
  await saveSettings(settings);
  refreshActionChoices(
    elements().actionParty.value,
    elements().actionStatus.value
  );
  renderSettings();
  showBanner(message, "success");
}


async function compressLogo(file: File): Promise<string> {
  if (!file.type.startsWith("image/")) {
    throw new Error("Choose a supported image file.");
  }

  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read the logo file."));
    reader.readAsDataURL(file);
  });

  const image = await new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not load the logo image."));
    img.src = dataUrl;
  });

  const maxW = 260;
  const maxH = 100;
  const scale = Math.min(maxW / image.width, maxH / image.height, 1);
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Could not prepare the logo image.");

  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(image, 0, 0, width, height);

  return canvas.toDataURL("image/png");
}

async function updatePartyLogo(partyId: string, file: File): Promise<void> {
  try {
    const logoDataUrl = await compressLogo(file);
    settings = {
      ...settings,
      parties: settings.parties.map((party) =>
        party.id === partyId ? { ...party, logoDataUrl } : party
      ),
    };
    await persistSettings("Party logo saved.");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}

async function removePartyLogo(partyId: string): Promise<void> {
  settings = {
    ...settings,
    parties: settings.parties.map((party) =>
      party.id === partyId ? { id: party.id, name: party.name } : party
    ),
  };

  try {
    await persistSettings("Party logo removed.");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}

async function removeParty(party: Party): Promise<void> {
  if (settings.parties.length <= 1) {
    showBanner("Keep at least one party in the library.", "error");
    return;
  }

  const inUse = currentIssue?.actions.some(
    (action) => action.party.toLowerCase() === party.name.toLowerCase()
  );
  if (inUse) {
    showBanner("This party is used by an existing action. Update that action before removing it.", "error");
    return;
  }

  settings = {
    ...settings,
    parties: settings.parties.filter((item) => item.id !== party.id),
  };

  try {
    await persistSettings("Party removed.");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}

function renderPartyLibrary(): void {
  const ui = elements();
  ui.partyLibraryList.innerHTML = "";

  settings.parties.forEach((party) => {
    const item = document.createElement("div");
    item.className = "library-item party-library-item";

    const preview = document.createElement("div");
    preview.className = "party-logo-preview";

    if (party.logoDataUrl) {
      const img = document.createElement("img");
      img.alt = `${party.name} logo`;
      img.src = party.logoDataUrl;
      preview.appendChild(img);
    } else {
      preview.textContent = party.name.slice(0, 3).toUpperCase();
    }

    const name = document.createElement("span");
    name.className = "library-name";
    name.textContent = party.name;

    const actions = document.createElement("div");
    actions.className = "party-item-actions";

    const upload = document.createElement("button");
    upload.className = "secondary";
    upload.textContent = party.logoDataUrl ? "Replace Logo" : "Add Logo";

    const fileInput = document.createElement("input");
    fileInput.className = "logo-file-input";
    fileInput.type = "file";
    fileInput.accept = "image/png,image/jpeg,image/webp";

    upload.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", () => {
      const file = fileInput.files?.[0];
      if (file) void updatePartyLogo(party.id, file);
      fileInput.value = "";
    });

    actions.append(upload, fileInput);

    if (party.logoDataUrl) {
      const removeLogo = document.createElement("button");
      removeLogo.className = "secondary";
      removeLogo.textContent = "Remove Logo";
      removeLogo.addEventListener("click", () => void removePartyLogo(party.id));
      actions.appendChild(removeLogo);
    }

    const remove = document.createElement("button");
    remove.className = "library-remove";
    remove.textContent = "Remove Party";
    remove.addEventListener("click", () => void removeParty(party));
    actions.appendChild(remove);

    item.append(preview, name, actions);
    ui.partyLibraryList.appendChild(item);
  });
}

function renderSettings(): void {
  const ui = elements();

  renderPartyLibrary();

  renderLibraryList(ui.statusLibraryList, settings.statuses, async (value) => {
    if (settings.statuses.length <= 1) {
      showBanner("Keep at least one status in the library.", "error");
      return;
    }

    const inUse = currentIssue?.actions.some(
      (action) => action.status.toLowerCase() === value.toLowerCase()
    );
    if (inUse) {
      showBanner("This status is used by an existing action. Update that action before removing it.", "error");
      return;
    }

    settings = {
      ...settings,
      statuses: settings.statuses.filter(
        (status) => status.toLowerCase() !== value.toLowerCase()
      ),
    };

    try {
      await persistSettings("Status removed.");
    } catch (error) {
      showBanner(error instanceof Error ? error.message : String(error), "error");
    }
  });
}

async function addParty(): Promise<void> {
  const ui = elements();
  const value = ui.newPartyName.value.trim();

  if (!value) {
    showBanner("Enter a party name first.", "error");
    return;
  }

  if (settings.parties.some((party) => party.name.toLowerCase() === value.toLowerCase())) {
    showBanner("That party already exists.", "error");
    return;
  }

  settings = {
    ...settings,
    parties: [...settings.parties, { id: `party-${Date.now().toString(36)}`, name: value }],
  };

  try {
    await persistSettings("Party added.");
    ui.newPartyName.value = "";
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}

async function addStatus(): Promise<void> {
  const ui = elements();
  const value = ui.newStatusName.value.trim();

  if (!value) {
    showBanner("Enter a status name first.", "error");
    return;
  }

  if (settings.statuses.some((status) => status.toLowerCase() === value.toLowerCase())) {
    showBanner("That status already exists.", "error");
    return;
  }

  settings = { ...settings, statuses: [...settings.statuses, value] };

  try {
    await persistSettings("Status added.");
    ui.newStatusName.value = "";
  } catch (error) {
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}

export async function initializeIssuePanel(): Promise<void> {
  const ui = elements();

  initializeTabs();
  refreshActionChoices();
  renderSettings();

  ui.save.addEventListener("click", saveOnly);
  ui.refresh.addEventListener("click", refreshSheet);
  ui.saveAction.addEventListener("click", saveAction);
  ui.cancelEditAction.addEventListener("click", resetActionEditor);
  ui.addParty.addEventListener("click", addParty);
  ui.addStatus.addEventListener("click", addStatus);

  try {
    const issue = await service.readSelectedIssue();
    populate(issue);
    if (issue) showBanner("Existing IssueFlow issue loaded.", "info");
  } catch (error) {
    populate(null);
    showBanner(error instanceof Error ? error.message : String(error), "error");
  }
}
