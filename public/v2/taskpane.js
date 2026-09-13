const TAG_APP = "PIT_APP";
const TAG_ISSUE_ID = "PIT_ISSUE_ID";
const TAG_CREATED_AT = "PIT_CREATED_AT";
const TAG_UPDATED_AT = "PIT_UPDATED_AT";
const TAG_DESCRIPTION_V2 = "PIT_DESCRIPTION_V2";
const TAG_AREA_CODE_V2 = "PIT_AREA_CODE_V2";
const TAG_ROOM_NAME_V2 = "PIT_ROOM_NAME_V2";
const TAG_ACTIONS_V2 = "PIT_ACTIONS_V2";
const TAG_PARTY_LIBRARY = "PIT_PARTY_LIBRARY";
const TAG_PARTY_LOGOS = "PIT_PARTY_LOGOS";
const TAG_STATUS_LIBRARY = "PIT_STATUS_LIBRARY";
const TAG_MANAGED = "PIT_MANAGED";
const TAG_MANAGED_ROLE = "PIT_MANAGED_ROLE";
const TAG_SUMMARY = "PIT_SUMMARY";
const TAG_SUMMARY_TYPE = "PIT_SUMMARY_TYPE";

// Legacy tags retained for migration/compatibility.
const TAG_BOX_TYPE = "PIT_BOX_TYPE";
const BOX_DESCRIPTION = "DESCRIPTION";

const DEFAULT_STATUSES = ["Open", "In Progress", "Pending", "Closed"];
const DEFAULT_PARTIES = ["Architect", "C&S", "MEP", "Main Contractor", "Subcontractor"];

const ui = {};
let editingActionId = null;
let pendingDeleteActionId = null;
let toastTimer = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  bindUi();
  initialize().catch(showError);
});

function bindUi() {
  [
    "issueId","setIssueId","areaCode","roomName","issueDescription","saveIssue","createIssueSheet",
    "currentIssueHeading","overallStatusChip","issueActionCount","issueCreatedDate","issueUpdatedDate",
    "issueNavigatorInput","issueNavigatorList","goToIssue",
    "actionParty","actionText","actionStatus","saveAction","clearActionForm",
    "actionEditorTitle","actionList","actionsCountBadge","actionsEmpty",
    "generateSummary","previewIssues","previewActions","previewOpen","previewClosed",
    "partyList","newParty","addParty","partyLogoFile","statusList","newStatus","addStatus","validatePresentation","toast"
  ].forEach(id => ui[id] = document.getElementById(id));

  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => activateTab(tab.dataset.tab));
  });

  ui.setIssueId.addEventListener("click", () => setIssueId().catch(showError));
  ui.saveIssue.addEventListener("click", () => saveIssueData().catch(showError));
  ui.createIssueSheet.addEventListener("click", () => createOrRefreshIssueSheet().catch(showError));
  ui.goToIssue.addEventListener("click", () => goToIssue().catch(showError));
  ui.issueNavigatorInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") goToIssue().catch(showError);
  });

  ui.saveAction.addEventListener("click", () => saveAction().catch(showError));
  ui.clearActionForm.addEventListener("click", clearActionForm);

  ui.generateSummary.addEventListener("click", () => generateSummary().catch(showError));
  ui.addParty.addEventListener("click", () => addParty().catch(showError));
  ui.partyLogoFile.addEventListener("change", () => handlePartyLogoSelected().catch(showError));
  ui.addStatus.addEventListener("click", () => addStatusValue().catch(showError));
  ui.validatePresentation.addEventListener("click", () => validatePresentation().catch(showError));
}

async function initialize() {
  await ensureLibraries();
  await refreshLibrariesUi();
  await refreshUiAfterChange();
}

function activateTab(name) {
  document.querySelectorAll(".tab").forEach(tab => tab.classList.toggle("active", tab.dataset.tab === name));
  document.querySelectorAll(".tab-panel").forEach(panel => panel.classList.toggle("active", panel.id === `panel-${name}`));
}

function cleanIssueId(value) {
  return String(value || "").trim().toUpperCase();
}

function cleanText(value) {
  return String(value || "").trim();
}

function formatDate(value) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  const datePart = date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  });
  const timePart = date.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  });
  return `${datePart} ${timePart}`;
}

function nowIso() {
  return new Date().toISOString();
}

function statusClass(status) {
  const s = cleanText(status).toLowerCase();
  if (s === "open") return "open";
  if (s === "in progress") return "in-progress";
  if (s === "pending") return "pending";
  if (s === "closed") return "closed";
  return "neutral";
}

const THEME = {
  navy: "#19364F",
  text: "#26343F",
  muted: "#6F7F8A",
  grid: "#D9E2E8",
  gridStrong: "#C4D0D8",
  white: "#FFFFFF"
};

function statusColor(status) {
  const s = cleanText(status).toLowerCase();
  if (s === "open") return { fill: "#FDEBEC", text: "#9D1C1C", line: "#E64A4A" };
  if (s === "in progress") return { fill: "#EAF3FF", text: "#174A7E", line: "#3A78C2" };
  if (s === "pending") return { fill: "#FFF5DB", text: "#7A560F", line: "#D89A2B" };
  if (s === "closed") return { fill: "#E9F7EE", text: "#1E603D", line: "#2F8F5B" };
  return { fill: "#F1F4F6", text: "#4B5964", line: "#AEBBC5" };
}

function computeOverallStatus(actions) {
  if (!actions.length) return "No actions";
  const statuses = actions.map(a => cleanText(a.status).toLowerCase());
  if (statuses.every(s => s === "closed")) return "Closed";
  if (statuses.includes("open")) return "Open";
  if (statuses.includes("in progress")) return "In Progress";
  if (statuses.includes("pending")) return "Pending";
  return actions[0].status || "Open";
}

async function getSelectedSlide(context) {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await context.sync();
  if (!selected.items.length) throw new Error("Select a PowerPoint slide first.");
  return selected.items[0];
}

async function getTagValue(tags, key, context) {
  const tag = tags.getItemOrNullObject(key);
  tag.load("value,isNullObject");
  await context.sync();
  return tag.isNullObject ? null : tag.value;
}

async function ensureLibraries() {
  await PowerPoint.run(async context => {
    const statusValue = await getTagValue(context.presentation.tags, TAG_STATUS_LIBRARY, context);
    if (!statusValue) context.presentation.tags.add(TAG_STATUS_LIBRARY, JSON.stringify(DEFAULT_STATUSES));

    const partyValue = await getTagValue(context.presentation.tags, TAG_PARTY_LIBRARY, context);
    if (!partyValue) context.presentation.tags.add(TAG_PARTY_LIBRARY, JSON.stringify(DEFAULT_PARTIES));

    await context.sync();
  });
}

async function readJsonPresentationTag(tagName, fallback) {
  return PowerPoint.run(async context => {
    const value = await getTagValue(context.presentation.tags, tagName, context);
    try {
      const parsed = JSON.parse(value || "[]");
      return Array.isArray(parsed) && parsed.length ? parsed : fallback.slice();
    } catch {
      return fallback.slice();
    }
  });
}

async function writeJsonPresentationTag(tagName, value) {
  await PowerPoint.run(async context => {
    context.presentation.tags.add(tagName, JSON.stringify(value));
    await context.sync();
  });
}

async function getStatuses() {
  return readJsonPresentationTag(TAG_STATUS_LIBRARY, DEFAULT_STATUSES);
}

async function getParties() {
  return readJsonPresentationTag(TAG_PARTY_LIBRARY, DEFAULT_PARTIES);
}


async function getPartyLogos() {
  return PowerPoint.run(async context => {
    const value = await getTagValue(context.presentation.tags, TAG_PARTY_LOGOS, context);
    try {
      const parsed = JSON.parse(value || "{}");
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch {
      return {};
    }
  });
}

async function savePartyLogos(logos) {
  await PowerPoint.run(async context => {
    context.presentation.tags.add(TAG_PARTY_LOGOS, JSON.stringify(logos || {}));
    await context.sync();
  });
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read the logo file."));
    reader.onload = () => resolve(String(reader.result || ""));
    reader.readAsDataURL(file);
  });
}

function loadBrowserImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onerror = () => reject(new Error("Could not open the selected logo image."));
    image.onload = () => resolve(image);
    image.src = dataUrl;
  });
}

async function optimizeLogo(file) {
  if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) {
    throw new Error("Use PNG, JPG or WebP for logos.");
  }

  const sourceUrl = await fileToDataUrl(file);
  const image = await loadBrowserImage(sourceUrl);
  const scale = Math.min(1, 320 / image.width, 120 / image.height);
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Image processing is unavailable.");

  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(image, 0, 0, width, height);

  let dataUrl = canvas.toDataURL("image/png");
  if (dataUrl.length > 180000) dataUrl = canvas.toDataURL("image/jpeg", 0.82);
  if (dataUrl.length > 220000) {
    throw new Error("Logo is too large after optimization. Please use a simpler or smaller image.");
  }
  return dataUrl;
}

function dataUrlToBase64(dataUrl) {
  const index = String(dataUrl || "").indexOf(",");
  return index >= 0 ? dataUrl.slice(index + 1) : "";
}

let pendingLogoParty = null;

async function choosePartyLogo(party) {
  pendingLogoParty = party;
  ui.partyLogoFile.value = "";
  ui.partyLogoFile.click();
}

async function handlePartyLogoSelected() {
  const file = ui.partyLogoFile.files?.[0];
  const party = pendingLogoParty;
  pendingLogoParty = null;
  if (!file || !party) return;

  const optimized = await optimizeLogo(file);
  const logos = await getPartyLogos();
  logos[party] = optimized;
  await savePartyLogos(logos);
  await refreshLibrariesUi();
  showToast(`Logo saved for ${party}.`);
}

async function removePartyLogo(party) {
  const logos = await getPartyLogos();
  if (!logos[party]) return;
  delete logos[party];
  await savePartyLogos(logos);
  await refreshLibrariesUi();
  showToast(`Logo removed for ${party}.`);
}


async function refreshLibrariesUi() {
  const [parties, statuses, logos] = await Promise.all([getParties(), getStatuses(), getPartyLogos()]);

  populateSelect(ui.actionParty, parties);
  populateSelect(ui.actionStatus, statuses);

  renderPartyLibraryList(ui.partyList, parties, logos);
  renderLibraryList(ui.statusList, statuses, value => removeStatus(value));
}

function populateSelect(select, values) {
  const previous = select.value;
  select.innerHTML = "";
  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (values.includes(previous)) select.value = previous;
}


function renderPartyLibraryList(container, parties, logos) {
  container.innerHTML = "";

  parties.forEach(party => {
    const row = document.createElement("div");
    row.className = "library-item";

    const main = document.createElement("div");
    main.className = "party-main";

    if (logos[party]) {
      const img = document.createElement("img");
      img.className = "party-logo-preview";
      img.src = logos[party];
      img.alt = `${party} logo`;
      main.appendChild(img);
    } else {
      const placeholder = document.createElement("div");
      placeholder.className = "party-logo-placeholder";
      placeholder.textContent = "LOGO";
      main.appendChild(placeholder);
    }

    const name = document.createElement("span");
    name.className = "party-name";
    name.textContent = party;
    main.appendChild(name);

    const actions = document.createElement("div");
    actions.className = "party-actions";

    const upload = document.createElement("button");
    upload.textContent = logos[party] ? "Replace" : "Upload";
    upload.addEventListener("click", () => choosePartyLogo(party).catch(showError));
    actions.appendChild(upload);

    if (logos[party]) {
      const removeLogo = document.createElement("button");
      removeLogo.textContent = "Logo ×";
      removeLogo.addEventListener("click", () => removePartyLogo(party).catch(showError));
      actions.appendChild(removeLogo);
    }

    const remove = document.createElement("button");
    remove.textContent = "Remove";
    remove.addEventListener("click", () => removeParty(party).catch(showError));
    actions.appendChild(remove);

    row.append(main, actions);
    container.appendChild(row);
  });
}

function renderLibraryList(container, values, removeHandler) {
  container.innerHTML = "";
  values.forEach(value => {
    const row = document.createElement("div");
    row.className = "library-item";
    const span = document.createElement("span");
    span.textContent = value;
    const button = document.createElement("button");
    button.textContent = "Remove";
    button.addEventListener("click", () => removeHandler(value).catch(showError));
    row.append(span, button);
    container.appendChild(row);
  });
}


async function refreshIssueNavigator() {
  const issues = await readAllIssues();

  ui.issueNavigatorList.innerHTML = "";
  issues.forEach(issue => {
    const option = document.createElement("option");
    option.value = issue.issueId;
    ui.issueNavigatorList.appendChild(option);
  });

  // Keep user-entered text. Only prefill when the box is empty.
  if (!cleanText(ui.issueNavigatorInput.value)) {
    const selected = await readSelectedIssue().catch(() => null);
    if (selected?.issueId) {
      ui.issueNavigatorInput.value = selected.issueId;
    }
  }
}

async function goToIssue() {
  const requestedId = cleanIssueId(ui.issueNavigatorInput.value);
  if (!requestedId) {
    throw new Error("Type or paste an Issue ID first.");
  }

  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const target = slides.items.find(slide => {
      const value = slide.tags.items.find(t => t.key === TAG_ISSUE_ID)?.value || "";
      return cleanIssueId(value) === requestedId;
    });

    if (!target) {
      throw new Error(`Issue ID ${requestedId} was not found.`);
    }

    context.presentation.setSelectedSlides([target.id]);
    await context.sync();
  });

  ui.issueNavigatorInput.value = requestedId;
  await loadCurrentIssueSafe();
  showToast(`Opened ${requestedId}.`);
}


async function refreshUiAfterChange() {
  await loadCurrentIssueSafe();
  await refreshIssueNavigator();
  await refreshSummaryPreview();
}

async function loadCurrentIssueSafe() {
  try {
    await loadCurrentIssue();
  } catch (error) {
    ui.issueId.value = "";
    ui.areaCode.value = "";
    ui.roomName.value = "";
    ui.issueDescription.value = "";
    updateCurrentIssueUi(null);
    renderActions([]);
  }
}

async function loadCurrentIssue() {
  const issue = await readSelectedIssue();
  ui.issueId.value = issue.issueId || "";
  ui.areaCode.value = issue.areaCode || "";
  ui.roomName.value = issue.roomName || "";
  ui.issueDescription.value = issue.description || "";
  updateCurrentIssueUi(issue);
  renderActions(issue.actions);
}

async function readSelectedIssue() {
  return PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    const descriptionTag = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context);
    const areaCode = await getTagValue(slide.tags, TAG_AREA_CODE_V2, context);
    const roomName = await getTagValue(slide.tags, TAG_ROOM_NAME_V2, context);
    const actionsTag = await getTagValue(slide.tags, TAG_ACTIONS_V2, context);
    const createdAt = await getTagValue(slide.tags, TAG_CREATED_AT, context);
    const updatedAt = await getTagValue(slide.tags, TAG_UPDATED_AT, context);

    let description = descriptionTag || "";
    let actions = [];
    try { actions = JSON.parse(actionsTag || "[]"); } catch { actions = []; }
    if (!Array.isArray(actions)) actions = [];

    // Migration helper: if V2 description is empty, try the V1 Description box.
    if (!description) {
      slide.shapes.load("items/id,items/tags/key,items/tags/value");
      await context.sync();
      const legacy = slide.shapes.items.find(shape => {
        const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
        return typeTag?.value === BOX_DESCRIPTION;
      });
      if (legacy) {
        legacy.textFrame.textRange.load("text");
        await context.sync();
        const text = legacy.textFrame.textRange.text || "";
        description = text.split(/\r?\n/).slice(1).join("\n").trim();
      }
    }

    return {
      slideId: slide.id,
      issueId: issueId || "",
      areaCode: areaCode || "",
      roomName: roomName || "",
      description,
      actions,
      createdAt: createdAt || "",
      updatedAt: updatedAt || ""
    };
  });
}

function updateCurrentIssueUi(issue) {
  const issueId = issue?.issueId || "";
  const actions = issue?.actions || [];
  const overall = computeOverallStatus(actions);

  ui.currentIssueHeading.textContent = issueId || "No issue selected";
  ui.issueActionCount.textContent = String(actions.length);
  ui.issueCreatedDate.textContent = formatDate(issue?.createdAt);
  ui.issueUpdatedDate.textContent = formatDate(issue?.updatedAt);

  ui.overallStatusChip.textContent = overall;
  ui.overallStatusChip.className = `chip ${statusClass(overall)}`;
}

async function setIssueId() {
  const issueId = cleanIssueId(ui.issueId.value);
  if (!issueId) throw new Error("Enter an Issue ID, for example MEP-001.");

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    for (const candidate of slides.items) {
      if (candidate.id === slide.id) continue;
      const tag = candidate.tags.items.find(t => t.key === TAG_ISSUE_ID);
      if (cleanIssueId(tag?.value) === issueId) {
        throw new Error(`${issueId} already exists on another slide.`);
      }
    }

    slide.tags.add(TAG_APP, "ISSUEFLOW_V2");
    slide.tags.add(TAG_ISSUE_ID, issueId);

    const created = await getTagValue(slide.tags, TAG_CREATED_AT, context);
    if (!created) slide.tags.add(TAG_CREATED_AT, nowIso());
    await context.sync();
  });

  await refreshUiAfterChange();
  ui.issueNavigatorInput.value = issueId;
  showToast(`Issue ID set to ${issueId}.`);
}

async function saveIssueData({ quiet = false } = {}) {
  const issueId = cleanIssueId(ui.issueId.value);
  if (!issueId) throw new Error("Set an Issue ID first.");
  const areaCode = cleanText(ui.areaCode.value).toUpperCase();
  const roomName = cleanText(ui.roomName.value);
  const description = cleanText(ui.issueDescription.value);

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const currentDescription = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context) || "";
    const currentAreaCode = await getTagValue(slide.tags, TAG_AREA_CODE_V2, context) || "";
    const currentRoomName = await getTagValue(slide.tags, TAG_ROOM_NAME_V2, context) || "";

    slide.tags.add(TAG_AREA_CODE_V2, areaCode);
    slide.tags.add(TAG_ROOM_NAME_V2, roomName);
    slide.tags.add(TAG_DESCRIPTION_V2, description);

    if (
      currentDescription !== description ||
      currentAreaCode !== areaCode ||
      currentRoomName !== roomName
    ) {
      slide.tags.add(TAG_UPDATED_AT, nowIso());
    }
    await context.sync();
  });

  await loadCurrentIssue();
  if (!quiet) showToast("Issue data saved.");
}

async function readActionsFromSelectedSlide() {
  const issue = await readSelectedIssue();
  if (!issue.issueId) throw new Error("Set an Issue ID first.");
  return issue;
}

function nextActionId(actions) {
  let max = 0;
  actions.forEach(action => {
    const match = /^ACT-(\d+)$/i.exec(action.id || "");
    if (match) max = Math.max(max, Number(match[1]));
  });
  return `ACT-${String(max + 1).padStart(3, "0")}`;
}

async function saveAction() {
  const issue = await readActionsFromSelectedSlide();
  const party = cleanText(ui.actionParty.value);
  const actionText = cleanText(ui.actionText.value);
  const status = cleanText(ui.actionStatus.value);

  if (!party) throw new Error("Choose a responsible party.");
  if (!actionText) throw new Error("Enter the required action.");
  if (!status) throw new Error("Choose a status.");

  const actions = issue.actions.slice();
  const timestamp = nowIso();
  const wasEditing = !!editingActionId;

  if (editingActionId) {
    const index = actions.findIndex(a => a.id === editingActionId);
    if (index < 0) throw new Error("The action being edited no longer exists.");
    actions[index] = {
      ...actions[index], party, action: actionText, status, updatedAt: timestamp
    };
  } else {
    actions.push({
      id: nextActionId(actions),
      party,
      action: actionText,
      status,

      createdAt: timestamp,
      updatedAt: ""
    });
  }

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.tags.add(TAG_ACTIONS_V2, JSON.stringify(actions));
    slide.tags.add(TAG_UPDATED_AT, timestamp);
    await context.sync();
  });

  clearActionForm();
  await refreshUiAfterChange();
  showToast(wasEditing ? "Action updated." : "Action added.");
}

function clearActionForm() {
  editingActionId = null;
  ui.actionEditorTitle.textContent = "Add action";
  ui.saveAction.textContent = "+ Add Action";
  ui.actionText.value = "";
}

function renderActions(actions) {
  ui.actionList.innerHTML = "";
  ui.actionsCountBadge.textContent = String(actions.length);
  ui.actionsEmpty.style.display = actions.length ? "none" : "block";

  actions.forEach(action => {
    const card = document.createElement("article");
    card.className = "action-card";

    const top = document.createElement("div");
    top.className = "action-top";
    const party = document.createElement("span");
    party.className = "action-party";
    party.textContent = action.party;
    const chip = document.createElement("span");
    chip.className = `chip ${statusClass(action.status)}`;
    chip.textContent = action.status;
    top.append(party, chip);

    const actionText = document.createElement("div");
    actionText.className = "action-text";
    actionText.textContent = action.action;


    const controls = document.createElement("div");
    controls.className = "action-controls";
    const edit = document.createElement("button");
    edit.textContent = "Edit";
    edit.addEventListener("click", () => editAction(action));
    const remove = document.createElement("button");
    remove.textContent = pendingDeleteActionId === action.id ? "Remove?" : "Remove";
    remove.className = "delete";
    remove.addEventListener("click", () => deleteAction(action.id).catch(showError));
    controls.append(edit, remove);

    card.append(top, actionText, controls);
    ui.actionList.appendChild(card);
  });
}

function editAction(action) {
  pendingDeleteActionId = null;
  editingActionId = action.id;
  ui.actionEditorTitle.textContent = `Edit ${action.id}`;
  ui.actionParty.value = action.party;
  ui.actionText.value = action.action;
  ui.actionStatus.value = action.status;
  ui.saveAction.textContent = "Save Changes";
  activateTab("actions");
}

async function deleteAction(actionId) {
  const issue = await readActionsFromSelectedSlide();
  const action = issue.actions.find(a => a.id === actionId);
  if (!action) return;

  // Office task panes may not support blocking browser confirmation dialogs.
  // Require a second click on Remove instead.
  if (pendingDeleteActionId !== actionId) {
    pendingDeleteActionId = actionId;
    showToast(`Click Remove again to delete ${action.party}.`);
    renderActions(issue.actions);
    return;
  }

  pendingDeleteActionId = null;

  const actions = issue.actions.filter(a => a.id !== actionId);
  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.tags.add(TAG_ACTIONS_V2, JSON.stringify(actions));
    slide.tags.add(TAG_UPDATED_AT, nowIso());
    await context.sync();
  });

  await refreshUiAfterChange();
  showToast("Action removed.");
}

async function addParty() {
  const value = cleanText(ui.newParty.value);
  if (!value) throw new Error("Enter a party name.");
  const parties = await getParties();
  if (parties.some(p => p.toLowerCase() === value.toLowerCase())) throw new Error("That party already exists.");
  parties.push(value);
  await writeJsonPresentationTag(TAG_PARTY_LIBRARY, parties);
  ui.newParty.value = "";
  await refreshLibrariesUi();
  showToast(`Party "${value}" added.`);
}

async function removeParty(value) {
  const parties = await getParties();
  if (parties.length <= 1) throw new Error("At least one party must remain.");
  const allIssues = await readAllIssues();
  const used = allIssues.some(issue => issue.actions.some(action => action.party === value));
  if (used) throw new Error(`"${value}" is already used by existing actions. Reassign those actions before removing it.`);
  await writeJsonPresentationTag(TAG_PARTY_LIBRARY, parties.filter(p => p !== value));
  await refreshLibrariesUi();
  showToast(`Party "${value}" removed.`);
}

async function addStatusValue() {
  const value = cleanText(ui.newStatus.value);
  if (!value) throw new Error("Enter a status name.");
  const statuses = await getStatuses();
  if (statuses.some(s => s.toLowerCase() === value.toLowerCase())) throw new Error("That status already exists.");
  statuses.push(value);
  await writeJsonPresentationTag(TAG_STATUS_LIBRARY, statuses);
  ui.newStatus.value = "";
  await refreshLibrariesUi();
  showToast(`Status "${value}" added.`);
}

async function removeStatus(value) {
  const statuses = await getStatuses();
  if (statuses.length <= 1) throw new Error("At least one status must remain.");
  const allIssues = await readAllIssues();
  const used = allIssues.some(issue => issue.actions.some(action => action.status === value));
  if (used) throw new Error(`"${value}" is already used by existing actions. Change those actions before removing it.`);
  await writeJsonPresentationTag(TAG_STATUS_LIBRARY, statuses.filter(s => s !== value));
  await refreshLibrariesUi();
  showToast(`Status "${value}" removed.`);
}

function markManaged(shape, role) {
  shape.tags.add(TAG_APP, "ISSUEFLOW_V2");
  shape.tags.add(TAG_MANAGED, "TRUE");
  shape.tags.add(TAG_MANAGED_ROLE, role);
  return shape;
}

function addText(slide, text, left, top, width, height, options = {}) {
  const box = slide.shapes.addTextBox(text, { left, top, width, height });
  box.textFrame.wordWrap = true;
  box.textFrame.textRange.font.name = options.font || "Aptos";
  box.textFrame.textRange.font.size = options.size || 10;
  box.textFrame.textRange.font.bold = !!options.bold;
  if (options.color) box.textFrame.textRange.font.color = options.color;
  return box;
}

function addRect(slide, left, top, width, height, fill, line, radius = false) {
  const rect = slide.shapes.addGeometricShape(
    radius ? PowerPoint.GeometricShapeType.roundRectangle : PowerPoint.GeometricShapeType.rectangle,
    { left, top, width, height }
  );
  rect.fill.setSolidColor(fill);
  rect.lineFormat.color = line;
  rect.lineFormat.weight = 1;
  return rect;
}


function addLine(slide, left, top, width, height, color = "#C9D5DE", weight = 1) {
  // Use a very thin rectangle as a divider instead of PowerPoint line shapes.
  // This is more predictable across PowerPoint hosts and avoids diagonal-line geometry issues.
  const horizontal = Math.abs(width) >= Math.abs(height);

  const options = horizontal
    ? {
        left: Math.min(left, left + width),
        top: top - (weight / 2),
        width: Math.max(0.5, Math.abs(width)),
        height: Math.max(0.5, weight)
      }
    : {
        left: left - (weight / 2),
        top: Math.min(top, top + height),
        width: Math.max(0.5, weight),
        height: Math.max(0.5, Math.abs(height))
      };

  const divider = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    options
  );
  divider.fill.setSolidColor(color);
  divider.lineFormat.color = color;
  divider.lineFormat.weight = 0;
  return divider;
}

function addSectionHeader(slide, title, left, top, width, options = {}) {
  const color = options.color || "#19364F";
  const lineColor = options.lineColor || "#8FB6CE";
  markManaged(addText(slide, title, left, top, width, 20, {
    size: options.size || 10,
    bold: true,
    color
  }), `${options.role || title}_TITLE`);
  markManaged(addLine(slide, left, top + 24, width, 0, lineColor, 1.4), `${options.role || title}_LINE`);
}

async function createOrRefreshIssueSheet() {
  await saveIssueData({ quiet: true });
  const issue = await readSelectedIssue();
  if (!issue.issueId) throw new Error("Set an Issue ID first.");

  const [parties, partyLogos] = await Promise.all([getParties(), getPartyLogos()]);

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    // Delete only IssueFlow-owned shapes.
    // Untagged/manual PowerPoint shapes (images, text, arrows, links, annotations)
    // must survive every refresh.
    for (const shape of slide.shapes.items) {
      const managed = shape.tags.items.find(t => t.key === TAG_MANAGED && t.value === "TRUE");
      if (managed) shape.delete();
    }
    await context.sync();

    const W = 960;
    const headerH = 76;
    const issueW = 132;
    const dateW = 176;
    const logoAreaW = W - issueW - dateW;
    const logoCellW = logoAreaW / Math.max(1, parties.length);

    markManaged(addRect(slide, 0, 0, logoAreaW, headerH, "#F8FBFD", "#F8FBFD"), "HEADER_LOGO_BG");

    parties.forEach((party, index) => {
      const x = index * logoCellW;
      if (index > 0) {
        markManaged(addLine(slide, x, 12, 0, headerH - 24, "#C7D9E5", 1), `HEADER_SEP_${index}`);
      }

      const logoData = partyLogos[party];
      let logoShown = false;

      if (logoData) {
        try {
          const base64 = dataUrlToBase64(logoData);
          const maxW = Math.max(28, logoCellW - 24);
          const logoW = Math.min(72, maxW);
          const logoBox = slide.shapes.addGeometricShape(
            PowerPoint.GeometricShapeType.rectangle,
            { left: x + (logoCellW - logoW) / 2, top: 9, width: logoW, height: 34 }
          );
          logoBox.fill.setImage(base64);
          logoBox.lineFormat.color = "#F8FBFD";
          logoBox.lineFormat.weight = 0;
          markManaged(logoBox, `HEADER_LOGO_${index}`);
          logoShown = true;
        } catch (error) {
          console.warn(`Could not render logo for ${party}.`, error);
        }
      }

      markManaged(addText(slide, party, x + 6, logoShown ? 48 : 30, logoCellW - 12, 16, {
        size: Math.max(6.5, Math.min(8.5, logoCellW / 13)),
        bold: true,
        color: THEME.navy
      }), `HEADER_PARTY_${index}`);
    });

    if (!parties.length) {
      markManaged(addText(slide, "Add parties / logos in Settings", 20, 29, logoAreaW - 40, 18, {
        size: 8.5, color: THEME.muted
      }), "HEADER_NO_PARTIES");
    }

    const issueX = logoAreaW;
    markManaged(addRect(slide, issueX, 0, issueW, headerH, "#FFF4F4", "#FFF4F4"), "ISSUE_BG");
    markManaged(addText(slide, "ISSUE ID", issueX + 14, 12, issueW - 28, 12, {
      size: 7.2, bold: true, color: "#6A4B4B"
    }), "ISSUE_LABEL");
    markManaged(addText(slide, issue.issueId, issueX + 14, 27, issueW - 28, 24, {
      size: 17, bold: true, color: "#17212B"
    }), "ISSUE_ID");
    if (issue.areaCode) {
      markManaged(addText(slide, issue.areaCode, issueX + 14, 54, issueW - 28, 12, {
        size: 7.3, bold: true, color: "#6A4B4B"
      }), "ISSUE_AREA");
    }

    const dateX = issueX + issueW;
    markManaged(addRect(slide, dateX, 0, dateW, headerH, "#FFF9E8", "#FFF9E8"), "DATE_BG");
    markManaged(addText(slide, `Created  ${formatDate(issue.createdAt)}`, dateX + 12, 18, dateW - 24, 16, {
      size: 7.3, bold: true, color: THEME.text
    }), "DATE_CREATED");
    if (issue.updatedAt) {
      markManaged(addText(slide, `Updated  ${formatDate(issue.updatedAt)}`, dateX + 12, 43, dateW - 24, 16, {
        size: 7.3, color: THEME.text
      }), "DATE_UPDATED");
    }

    markManaged(addRect(slide, 0, headerH, W, 2, THEME.navy, THEME.navy), "HEADER_BASE");

    const mainTop = 96;
    const sectionH = 28;
    const bodyTop = mainTop + sectionH;
    const bodyBottom = 520;
    const bodyH = bodyBottom - bodyTop;

    const margin = 18;
    const gap = 10;
    const descW = 246;
    const roomW = 38;
    const refW = 300;
    const actionsW = W - margin * 2 - descW - roomW - refW - gap * 3;

    const descX = margin;
    const roomX = descX + descW + gap;
    const refX = roomX + roomW + gap;
    const actionsX = refX + refW + gap;

    function sectionHeader(title, x, width, role) {
      markManaged(addRect(slide, x, mainTop, width, sectionH, THEME.navy, THEME.navy), `${role}_HDR_BG`);
      markManaged(addText(slide, title, x + 10, mainTop + 8, width - 20, 14, {
        size: 8.8, bold: true, color: "#FFFFFF"
      }), `${role}_HDR_TEXT`);
    }

    sectionHeader("ISSUE DESCRIPTION", descX, descW, "DESC");
    sectionHeader("REFERENCE IMAGES", refX, refW, "REF");
    sectionHeader("ACTIONS", actionsX, actionsW, "ACT");

    markManaged(addRect(slide, descX, bodyTop, descW, bodyH, "#FFFFFF", THEME.grid), "DESC_BODY");
    markManaged(addText(slide, issue.description || "No description yet.", descX + 12, bodyTop + 14, descW - 24, bodyH - 28, {
      size: 10.2, color: THEME.text
    }), "DESC_TEXT");

    // Area / Room strip: 50/50 vertical split.
    markManaged(addRect(slide, roomX, bodyTop, roomW, bodyH, "#FFFFFF", THEME.grid), "ROOM_BODY");

    const halfH = bodyH / 2;
    markManaged(
      addLine(slide, roomX, bodyTop + halfH, roomW, 0, THEME.gridStrong, 1),
      "ROOM_HALF_DIVIDER"
    );

    // Rotated textboxes must be positioned by their CENTER, not by their
    // unrotated left edge. Otherwise the textbox rotates out of the narrow strip.
    const stripCenterX = roomX + (roomW / 2);
    const rotatedBoxW = Math.max(60, halfH - 34);
    const rotatedBoxH = 20;

    const areaCenterY = bodyTop + (halfH / 2);
    const areaLabel = markManaged(addText(
      slide,
      issue.areaCode || "Area",
      stripCenterX - (rotatedBoxW / 2),
      areaCenterY - (rotatedBoxH / 2),
      rotatedBoxW,
      rotatedBoxH,
      {
        size: 9.5,
        bold: true,
        color: THEME.text
      }
    ), "ROOM_AREA_TEXT");
    areaLabel.rotation = 270;

    const roomCenterY = bodyTop + halfH + (halfH / 2);
    const roomLabel = markManaged(addText(
      slide,
      issue.roomName || "Room / Space",
      stripCenterX - (rotatedBoxW / 2),
      roomCenterY - (rotatedBoxH / 2),
      rotatedBoxW,
      rotatedBoxH,
      {
        size: 9.5,
        bold: true,
        color: THEME.text
      }
    ), "ROOM_TEXT");
    roomLabel.rotation = 270;

    // Reference Images is user-owned/manual content.
    // Do NOT create a filled body shape here because a newly-created fill would
    // sit above manually pasted images/text after refresh.
    // Only draw a thin border around the manual area.
    markManaged(addLine(slide, refX, bodyTop, refW, 0, THEME.grid, 1), "REF_TOP");
    markManaged(addLine(slide, refX, bodyBottom, refW, 0, THEME.grid, 1), "REF_BOTTOM");
    markManaged(addLine(slide, refX, bodyTop, 0, bodyH, THEME.grid, 1), "REF_LEFT");
    markManaged(addLine(slide, refX + refW, bodyTop, 0, bodyH, THEME.grid, 1), "REF_RIGHT");

    const colHeadH = 30;
    const rowsTop = bodyTop + colHeadH;
    const rowsH = bodyBottom - rowsTop;
    const partyW = Math.round(actionsW * 0.24);
    const statusW = Math.round(actionsW * 0.23);
    const actionW = actionsW - partyW - statusW;

    const partyX = actionsX;
    const actionX = partyX + partyW;
    const statusX = actionX + actionW;

    markManaged(addRect(slide, actionsX, bodyTop, actionsW, colHeadH, "#F1F5F8", THEME.gridStrong), "ACT_COL_HDR");
    markManaged(addText(slide, "Action By", partyX + 8, bodyTop + 9, partyW - 16, 13, {
      size: 8.2, bold: true, color: THEME.navy
    }), "ACT_HDR_PARTY");
    markManaged(addText(slide, "Action Required", actionX + 8, bodyTop + 9, actionW - 16, 13, {
      size: 8.2, bold: true, color: THEME.navy
    }), "ACT_HDR_ACTION");
    markManaged(addText(slide, "Status", statusX + 8, bodyTop + 9, statusW - 16, 13, {
      size: 8.2, bold: true, color: THEME.navy
    }), "ACT_HDR_STATUS");

    markManaged(addRect(slide, actionsX, bodyTop, actionsW, bodyH, "#FFFFFF", THEME.grid), "ACT_BODY");
    markManaged(addLine(slide, actionX, bodyTop, 0, bodyH, THEME.gridStrong, 1), "ACT_SEP1");
    markManaged(addLine(slide, statusX, bodyTop, 0, bodyH, THEME.gridStrong, 1), "ACT_SEP2");
    markManaged(addLine(slide, actionsX, bodyTop + colHeadH, actionsW, 0, THEME.gridStrong, 1), "ACT_HEAD_LINE");

    const actions = issue.actions || [];
    const overall = computeOverallStatus(actions);
    const overallColor = statusColor(overall);
    markManaged(addText(slide, `Overall: ${overall}`, actionsX + actionsW - 120, mainTop + 8, 110, 14, {
      size: 7.4, bold: true, color: overallColor.text
    }), "ACT_OVERALL");

    if (!actions.length) {
      markManaged(addText(slide, "No actions yet", actionsX + 12, rowsTop + 22, actionsW - 24, 18, {
        size: 9, color: THEME.muted
      }), "ACT_EMPTY");
    } else {
      const rowH = rowsH / actions.length;

      actions.forEach((action, index) => {
        const y = rowsTop + index * rowH;
        const sc = statusColor(action.status);
        const fontSize = rowH >= 72 ? 9.2 : rowH >= 50 ? 8.5 : 7.4;

        if (index > 0) {
          markManaged(addLine(slide, actionsX, y, actionsW, 0, THEME.grid, 1), `ACT_ROW_${index}`);
        }

        markManaged(addText(slide, action.party || "—", partyX + 10, y + 12, partyW - 20, Math.max(20, rowH - 22), {
          size: fontSize, bold: true, color: THEME.text
        }), `ACT_PARTY_${index}`);

        markManaged(addText(slide, action.action || "—", actionX + 10, y + 12, actionW - 20, Math.max(20, rowH - 22), {
          size: fontSize, color: THEME.text
        }), `ACT_TEXT_${index}`);

        markManaged(addRect(slide, statusX, y, statusW, rowH, sc.fill, sc.line), `ACT_STATUS_BG_${index}`);
        markManaged(addText(slide, action.status || "—", statusX + 10, y + 12, statusW - 20, Math.max(18, rowH - 22), {
          size: fontSize, bold: true, color: sc.text
        }), `ACT_STATUS_${index}`);
      });
    }

    await context.sync();
  });

  showToast("Issue sheet refreshed. Manual reference images and annotations were preserved.");
}

async function readAllIssues() {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value,items/shapes/items/id,items/shapes/items/tags/key,items/shapes/items/tags/value");
    await context.sync();

    const issues = [];

    for (const slide of slides.items) {
      const issueId = slide.tags.items.find(t => t.key === TAG_ISSUE_ID)?.value || "";
      if (!issueId) continue;

      let areaCode = slide.tags.items.find(t => t.key === TAG_AREA_CODE_V2)?.value || "";
      let roomName = slide.tags.items.find(t => t.key === TAG_ROOM_NAME_V2)?.value || "";
      let description = slide.tags.items.find(t => t.key === TAG_DESCRIPTION_V2)?.value || "";
      const createdAt = slide.tags.items.find(t => t.key === TAG_CREATED_AT)?.value || "";
      const updatedAt = slide.tags.items.find(t => t.key === TAG_UPDATED_AT)?.value || "";

      let actions = [];
      const actionJson = slide.tags.items.find(t => t.key === TAG_ACTIONS_V2)?.value || "[]";
      try {
        const parsed = JSON.parse(actionJson);
        if (Array.isArray(parsed)) actions = parsed;
      } catch {
        actions = [];
      }

      // Fallback migration: recover the visible V2 managed text if a tag is missing.
      if (!description || !roomName) {
        const descShape = slide.shapes.items.find(shape =>
          shape.tags.items.some(t => t.key === TAG_MANAGED_ROLE && t.value === "DESC_TEXT")
        );
        const roomShape = slide.shapes.items.find(shape =>
          shape.tags.items.some(t => t.key === TAG_MANAGED_ROLE && t.value === "ROOM_TEXT")
        );

        if (descShape) descShape.textFrame.textRange.load("text");
        if (roomShape) roomShape.textFrame.textRange.load("text");
        await context.sync();

        if (!description && descShape) description = cleanText(descShape.textFrame.textRange.text);
        if (!roomName && roomShape) roomName = cleanText(roomShape.textFrame.textRange.text);
      }

      // Area code is also visible in the Issue ID header in newer sheets.
      if (!areaCode) {
        const areaShape = slide.shapes.items.find(shape =>
          shape.tags.items.some(t => t.key === TAG_MANAGED_ROLE && (t.value === "ISSUE_AREA" || t.value === "ROOM_AREA_TEXT"))
        );
        if (areaShape) {
          areaShape.textFrame.textRange.load("text");
          await context.sync();
          areaCode = cleanText(areaShape.textFrame.textRange.text);
        }
      }

      issues.push({
        slideId: slide.id,
        issueId,
        areaCode,
        roomName,
        description,
        actions,
        createdAt,
        updatedAt
      });
    }

    issues.sort((a, b) =>
      a.issueId.localeCompare(b.issueId, undefined, { numeric: true, sensitivity: "base" })
    );

    return issues;
  });
}


async function deleteExistingSummarySlides() {
  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const summarySlideIds = slides.items
      .filter(slide => slide.tags.items.some(t => t.key === TAG_SUMMARY && t.value === "TRUE"))
      .map(slide => slide.id);

    summarySlideIds.forEach(id => slides.getItem(id).delete());
    await context.sync();
  });
}

async function addCleanGeneratedSlide(context, type) {
  const slides = context.presentation.slides;
  const count = slides.getCount();
  slides.add();
  await context.sync();

  const slide = slides.getItemAt(count.value);
  slide.load("id,shapes/items/id");
  await context.sync();

  // Remove placeholders inherited from the current PowerPoint layout.
  for (const shape of slide.shapes.items) {
    shape.delete();
  }
  await context.sync();

  slide.tags.add(TAG_SUMMARY, "TRUE");
  slide.tags.add(TAG_SUMMARY_TYPE, type);
  slide.tags.add(TAG_APP, "ISSUEFLOW_V2");
  await context.sync();

  return slide;
}

function buildDashboard(slide, issues) {
  const allActions = issues.flatMap(issue => issue.actions);
  const statusCounts = {};
  allActions.forEach(action => statusCounts[action.status] = (statusCounts[action.status] || 0) + 1);
  const openIssues = issues.filter(issue => computeOverallStatus(issue.actions) !== "Closed").length;
  const closedIssues = issues.length - openIssues;

  addRect(slide, 0, 0, 960, 62, "#13283A", "#13283A");
  addText(slide, "ISSUEFLOW DASHBOARD", 36, 17, 460, 26, { size: 22, bold: true, color: "#FFFFFF" });
  addText(slide, `Refreshed ${new Date().toLocaleString("en-GB")}`, 690, 20, 230, 17, { size: 8.5, color: "#DCE7EF" });

  const cards = [
    ["Issues", issues.length, "#445B70"],
    ["Open issues", openIssues, "#D64545"],
    ["Closed issues", closedIssues, "#2F8F5B"],
    ["Actions", allActions.length, "#3A78C2"]
  ];
  cards.forEach((card, i) => {
    const left = 40 + i * 220;
    addRect(slide, left, 88, 195, 78, "#FFFFFF", "#DCE4EA", true);
    addRect(slide, left, 88, 6, 78, card[2], card[2]);
    addText(slide, String(card[1]), left + 20, 101, 150, 26, { size: 22, bold: true, color: "#17212B" });
    addText(slide, card[0], left + 20, 135, 150, 16, { size: 9, color: "#68747F" });
  });

  addText(slide, "Action status", 40, 205, 300, 20, { size: 14, bold: true, color: "#17212B" });
  const max = Math.max(1, ...Object.values(statusCounts), 1);
  const statuses = Object.entries(statusCounts);
  statuses.forEach(([status, count], i) => {
    const top = 245 + i * 46;
    const sc = statusColor(status);
    addText(slide, status, 48, top, 130, 18, { size: 10, color: "#3B4853" });
    addRect(slide, 180, top + 2, 500, 15, "#EDF1F4", "#EDF1F4", true);
    addRect(slide, 180, top + 2, Math.max(6, 500 * count / max), 15, sc.line, sc.line, true);
    addText(slide, String(count), 695, top - 1, 50, 18, { size: 10, bold: true, color: "#3B4853" });
  });

  // Actions by party: top 6.
  const partyCounts = {};
  allActions.forEach(action => partyCounts[action.party] = (partyCounts[action.party] || 0) + 1);
  const parties = Object.entries(partyCounts).sort((a,b) => b[1]-a[1]).slice(0,6);
  addText(slide, "Actions by party", 760, 205, 160, 20, { size: 14, bold: true, color: "#17212B" });
  parties.forEach(([party, count], i) => {
    addText(slide, `${party}`, 760, 245 + i*34, 125, 16, { size: 9, color: "#4B5964" });
    addText(slide, `${count}`, 890, 245 + i*34, 30, 16, { size: 9, bold: true, color: "#1F78B4" });
  });

}

function groupIssuesForRegister(issues, maxRows = 9) {
  const pages = [];
  let page = [];
  let used = 0;

  for (const issue of issues) {
    const rows = Math.max(1, issue.actions.length);
    if (rows <= maxRows && used > 0 && used + rows > maxRows) {
      pages.push(page);
      page = [];
      used = 0;
    }

    if (rows <= maxRows) {
      page.push({ issue, start: 0, count: rows });
      used += rows;
    } else {
      let start = 0;
      while (start < rows) {
        if (used === maxRows) {
          pages.push(page); page = []; used = 0;
        }
        const available = maxRows - used;
        const count = Math.min(available, rows - start);
        page.push({ issue, start, count });
        used += count;
        start += count;
        if (used === maxRows) {
          pages.push(page); page = []; used = 0;
        }
      }
    }
  }
  if (page.length) pages.push(page);
  return pages.length ? pages : [[]];
}

function buildRegisterPage(slide, groups, pageNo, pageCount) {
  addRect(slide, 0, 0, 960, 58, THEME.navy, THEME.navy);
  addText(slide, "ACTION REGISTER", 34, 16, 430, 24, {
    size: 21, bold: true, color: "#FFFFFF"
  });
  addText(slide, `Page ${pageNo} of ${pageCount}`, 800, 19, 120, 16, {
    size: 8.5, color: "#DCE7EF"
  });

  const left = 18;
  const top = 82;
  const headH = 30;
  const rowH = 42;
  const widths = [92, 104, 214, 104, 281, 125];
  const headers = ["Issue ID", "Area / Room", "Description", "Action By", "Action Required", "Status"];

  let x = left;
  headers.forEach((header, index) => {
    addRect(slide, x, top, widths[index], headH, THEME.navy, THEME.navy);
    addText(slide, header, x + 8, top + 9, widths[index] - 16, 12, {
      size: 8.0, bold: true, color: "#FFFFFF"
    });
    x += widths[index];
  });

  let globalRow = 0;

  groups.forEach(group => {
    const issue = group.issue;
    const actions = issue.actions.length ? issue.actions : [{ party: "—", action: "No actions", status: "—" }];
    const visibleActions = actions.slice(group.start, group.start + group.count);
    const groupTop = top + headH + globalRow * rowH;
    const groupHeight = visibleActions.length * rowH;
    const issueFill = globalRow % 2 === 0 ? "#F7F9FB" : "#FFFFFF";

    addRect(slide, left, groupTop, widths[0], groupHeight, issueFill, THEME.grid);
    addText(slide, issue.issueId, left + 8, groupTop + 11, widths[0] - 16, Math.max(18, groupHeight - 16), {
      size: 9.0, bold: true, color: "#1769AA"
    });

    const areaLeft = left + widths[0];
    addRect(slide, areaLeft, groupTop, widths[1], groupHeight, issueFill, THEME.grid);
    addText(slide, [issue.areaCode, issue.roomName].filter(Boolean).join("\n") || "—",
      areaLeft + 8, groupTop + 10, widths[1] - 16, Math.max(18, groupHeight - 16), {
        size: 8.1, color: THEME.text
      });

    const descLeft = areaLeft + widths[1];
    addRect(slide, descLeft, groupTop, widths[2], groupHeight, issueFill, THEME.grid);
    addText(slide, issue.description || "—", descLeft + 8, groupTop + 10, widths[2] - 16, Math.max(18, groupHeight - 16), {
      size: 8.1, color: THEME.text
    });

    visibleActions.forEach((action, localIndex) => {
      const y = groupTop + localIndex * rowH;
      const rowFill = (globalRow + localIndex) % 2 === 0 ? "#FFFFFF" : "#FAFBFC";
      let dx = descLeft + widths[2];

      addRect(slide, dx, y, widths[3], rowH, rowFill, THEME.grid);
      addText(slide, action.party || "—", dx + 8, y + 11, widths[3] - 16, rowH - 14, {
        size: 8.1, color: THEME.text
      });
      dx += widths[3];

      addRect(slide, dx, y, widths[4], rowH, rowFill, THEME.grid);
      addText(slide, action.action || "—", dx + 8, y + 11, widths[4] - 16, rowH - 14, {
        size: 8.1, color: THEME.text
      });
      dx += widths[4];

      const sc = statusColor(action.status);
      addRect(slide, dx, y, widths[5], rowH, sc.fill, sc.line);
      addText(slide, action.status || "—", dx + 10, y + 11, widths[5] - 20, rowH - 14, {
        size: 8.3, bold: true, color: sc.text
      });

      globalRow++;
    });
  });
}

async function generateSummary() {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new Error("Summary generation requires PowerPointApi 1.8 or later.");
  }

  const issues = await readAllIssues();
  if (!issues.length) throw new Error("No tracked issues found.");

  await deleteExistingSummarySlides();
  const pages = groupIssuesForRegister(issues, 9);
  let dashboardId = "";
  const registerIds = [];

  await PowerPoint.run(async context => {
    const dashboard = await addCleanGeneratedSlide(context, "DASHBOARD");
    dashboardId = dashboard.id;
    for (let i = 0; i < pages.length; i++) {
      const register = await addCleanGeneratedSlide(context, "REGISTER");
      registerIds.push(register.id);
    }
  });

  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    const generated = [dashboardId, ...registerIds];
    for (let i = generated.length - 1; i >= 0; i--) slides.getItem(generated[i]).moveTo(0);
    await context.sync();
  });

  await PowerPoint.run(async context => {
    const dashboard = context.presentation.slides.getItem(dashboardId);
    buildDashboard(dashboard, issues);
    pages.forEach((groups, index) => {
      const register = context.presentation.slides.getItem(registerIds[index]);
      buildRegisterPage(register, groups, index + 1, pages.length);
    });
    await context.sync();
  });

  await refreshSummaryPreview();
  showToast(`Summary refreshed: 1 dashboard + ${pages.length} register page(s).`);
}

async function refreshSummaryPreview() {
  const issues = await readAllIssues();
  const actions = issues.flatMap(issue => issue.actions);
  ui.previewIssues.textContent = String(issues.length);
  ui.previewActions.textContent = String(actions.length);
  ui.previewOpen.textContent = String(actions.filter(a => cleanText(a.status).toLowerCase() !== "closed").length);
  ui.previewClosed.textContent = String(actions.filter(a => cleanText(a.status).toLowerCase() === "closed").length);
}

async function validatePresentation() {
  const issues = await readAllIssues();
  const parties = await getParties();
  const statuses = await getStatuses();
  const errors = [];
  const seen = new Set();

  issues.forEach(issue => {
    const id = cleanIssueId(issue.issueId);
    if (!id) errors.push("A tracked slide has no Issue ID.");
    if (seen.has(id)) errors.push(`${id}: duplicate Issue ID.`);
    seen.add(id);
    if (!cleanText(issue.areaCode)) errors.push(`${id}: area code is empty.`);
    if (!cleanText(issue.roomName)) errors.push(`${id}: room / space is empty.`);
    if (!cleanText(issue.description)) errors.push(`${id}: description is empty.`);
    issue.actions.forEach(action => {
      if (!cleanText(action.party)) errors.push(`${id}/${action.id}: party is empty.`);
      else if (!parties.includes(action.party)) errors.push(`${id}/${action.id}: party "${action.party}" is not in the party library.`);
      if (!cleanText(action.action)) errors.push(`${id}/${action.id}: action text is empty.`);
      if (!cleanText(action.status)) errors.push(`${id}/${action.id}: status is empty.`);
      else if (!statuses.includes(action.status)) errors.push(`${id}/${action.id}: status "${action.status}" is not in the status library.`);
    });
  });

  if (!issues.length) return showToast("No tracked issues found.", true);
  if (!errors.length) return showToast(`Validation passed: ${issues.length} issue(s) checked.`);
  showToast(`Validation found ${errors.length} problem(s). Open the console for details.`, true);
  console.warn("IssueFlow validation", errors);
}

function showToast(message, isError = false) {
  clearTimeout(toastTimer);
  ui.toast.textContent = message;
  ui.toast.className = `toast show${isError ? " error" : ""}`;
  toastTimer = setTimeout(() => ui.toast.className = "toast", 4200);
}

function showError(error) {
  console.error(error);
  showToast(error?.message || String(error), true);
}
